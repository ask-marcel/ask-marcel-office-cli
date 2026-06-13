/**
 * Ask Marcel Companion v3.2 - Background Service Worker
 * 
 * 通过 chrome.debugger 监听 OAuth 服务器的响应，
 * 从中提取多种 token（Graph、chatsvcagg、IC3、elevated）并发送给 CLI。
 * 同时捕获 chatsvcagg region（从 /api/csa/<region>/ URL 中提取）。
 */

console.log('[Ask Marcel] ✓ Service Worker 已启动 v3.2');

let callbackPort = null;
let debuggedTabId = null;
let tokenSent = false;

// 跟踪待处理的请求 ID
let pendingRequestIds = new Set();

// 多 token 捕获常量
const BASIC_TEAMS_APP_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const ELEVATED_APP_IDS = [
  'c0ab8ce9-e9a0-42e7-b064-33d422df41f1', // M365ChatClient
  '4765445b-32c6-49b0-83e6-1d93765276ca', // OfficeHome
];
const CHATSVCAGG_AUD = 'https://chatsvcagg.teams.microsoft.com';
const IC3_AUD = 'https://ic3.teams.office.com';
const GRAPH_AUD = 'https://graph.microsoft.com';
const GRAPH_AUD_ALT = '00000003-0000-0000-c000-000000000000';

// 多 token 捕获状态
let capturedGraph = null;
let capturedChatsvcagg = null;
let capturedIc3 = null;
let capturedElevated = null;
let capturedRegion = null;
let sendTimer = null;

// 每次捕获到新 token 后等待多久再发送（ms）
// Teams 页面加载后 MSAL 会依次静默获取 chatsvcagg/IC3 token
const SEND_DELAY_MS = 5000;

// 使用 webRequest.onBeforeRequest 拦截请求，捕获 ask_marcel_port
chrome.webRequest.onBeforeRequest.addListener(
  async (details) => {
    if (tokenSent) return;
    if (details.tabId <= 0) return;

    console.log('[Ask Marcel] webRequest 检测到请求: ' + details.url.substring(0, 120));

    try {
      const url = new URL(details.url);
      if (url.hostname !== 'teams.microsoft.com' && url.hostname !== 'teams.live.com') return;

      const port = url.searchParams.get('ask_marcel_port');
      if (!port) return;

      console.log('[Ask Marcel] ✓ webRequest 捕获到 ask_marcel_port=' + port + ', tab=' + details.tabId);
      await handlePortCapture(port, details.tabId);
    } catch (e) {
      console.error('[Ask Marcel] webRequest 错误:', e.message);
    }
  },
  { urls: ['*://teams.microsoft.com/*', '*://teams.live.com/*'] },
  []
);

async function handlePortCapture(port, tabId) {
  if (callbackPort && debuggedTabId === tabId) return;

  callbackPort = port;
  debuggedTabId = tabId;

  await attachDebugger(tabId);
}

async function attachDebugger(tabId) {
  try {
    await chrome.debugger.attach({ tabId }, '1.3');
    console.log('[Ask Marcel] ✓ Debugger attached to tab ' + tabId);

    await chrome.debugger.sendCommand({ tabId }, 'Network.enable', {});
    console.log('[Ask Marcel] ✓ Network monitoring enabled, waiting for tokens...');
  } catch (e) {
    console.error('[Ask Marcel] ✗ Debugger attach failed:', e.message);
  }
}

// 检查是否是 OAuth 服务器
function isOAuthHost(url) {
  return url.includes('login.microsoftonline.com') ||
         url.includes('login.live.com') ||
         url.includes('login.microsoft.com');
}

// 从 URL 中提取 chatsvcagg region
function parseRegion(url) {
  const match = /^https:\/\/teams\.microsoft\.com\/api\/(?:csa|chatsvc)\/([a-z0-9-]+)\//i.exec(url);
  return match ? match[1].toLowerCase() : null;
}

// 安全地解析 JWT payload
function decodeJwtPayload(token) {
  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;
    let payload = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    while (payload.length % 4) payload += '=';
    return JSON.parse(atob(payload));
  } catch (e) {
    return null;
  }
}

// 每次捕获到新 token 时调用：重置发送计时器
// 这样后续到达的 token 也能被收集到
function scheduleSend() {
  if (tokenSent) return;
  if (sendTimer) clearTimeout(sendTimer);
  sendTimer = setTimeout(() => {
    sendTimer = null;
    sendAllTokens();
  }, SEND_DELAY_MS);
}

// 发送所有已捕获的 token 到 CLI
async function sendAllTokens() {
  if (tokenSent || !callbackPort) return;
  if (!capturedGraph) {
    console.log('[Ask Marcel] 没有 Graph token，无法发送');
    return;
  }

  tokenSent = true;

  const url = 'http://127.0.0.1:' + callbackPort + '/token';
  const payload = {
    access_token: capturedGraph,
  };

  if (capturedChatsvcagg) {
    payload.chatsvcagg_access_token = capturedChatsvcagg;
    console.log('[Ask Marcel] ✓ 包含 chatsvcagg token');
  }
  if (capturedIc3) {
    payload.ic3_access_token = capturedIc3;
    console.log('[Ask Marcel] ✓ 包含 IC3 token');
  }
  if (capturedElevated) {
    payload.elevated_access_token = capturedElevated;
    console.log('[Ask Marcel] ✓ 包含 elevated token');
  }
  if (capturedRegion) {
    payload.chatsvcagg_region = capturedRegion;
    console.log('[Ask Marcel] ✓ 包含 region: ' + capturedRegion);
  }

  const tokenCount = 1 + (capturedChatsvcagg ? 1 : 0) + (capturedIc3 ? 1 : 0) + (capturedElevated ? 1 : 0);
  console.log('[Ask Marcel] 共捕获 ' + tokenCount + ' 个 token，发送中...');

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (response.ok) {
      console.log('[Ask Marcel] ✓ Tokens 发送成功！');
    } else {
      console.error('[Ask Marcel] ✗ 发送失败:', response.status);
      tokenSent = false;
    }
  } catch (e) {
    console.error('[Ask Marcel] ✗ 发送错误:', e.message);
    tokenSent = false;
  }

  setTimeout(async () => {
    if (debuggedTabId !== null) {
      try {
        await chrome.debugger.detach({ tabId: debuggedTabId });
        console.log('[Ask Marcel] ✓ Debugger detached (清理完成)');
      } catch (e) {}
      debuggedTabId = null;
    }
  }, 2000);
}

// 监听 debugger 事件
chrome.debugger.onEvent.addListener((source, method, params) => {
  if (tokenSent || !callbackPort) return;
  if (source.tabId !== debuggedTabId) return;

  // 监听 chatsvcagg region URL（从请求 URL 中提取）
  if (method === 'Network.requestWillBeSent') {
    const url = params.request?.url;
    if (url && !capturedRegion) {
      const region = parseRegion(url);
      if (region) {
        capturedRegion = region;
        console.log('[Ask Marcel] ✓ 捕获到 region: ' + region);
      }
    }
  }

  // 监听 OAuth 响应，记录 requestId 等待 loadingFinished
  if (method === 'Network.responseReceived') {
    const url = params.response?.url;
    if (!url) return;

    // 检查是否是 OAuth 服务器的 JSON 响应
    if (!isOAuthHost(url)) return;
    
    const contentType = params.response?.headers?.['content-type'] || 
                        params.response?.headers?.['Content-Type'] || '';
    if (!contentType.includes('json')) return;

    console.log('[Ask Marcel] 检测到 OAuth 响应: ' + url.substring(0, 80));
    
    pendingRequestIds.add(params.requestId);
    console.log('[Ask Marcel] 等待请求完成: ' + params.requestId);
  }

  // 响应完全加载后获取 body
  if (method === 'Network.loadingFinished') {
    const requestId = params.requestId;
    if (!pendingRequestIds.has(requestId)) return;

    pendingRequestIds.delete(requestId);
    console.log('[Ask Marcel] 请求完成，获取响应体: ' + requestId);

    chrome.debugger.sendCommand(
      { tabId: debuggedTabId },
      'Network.getResponseBody',
      { requestId: requestId },
      (result) => {
        if (chrome.runtime.lastError || !result) {
          console.log('[Ask Marcel] 获取响应体失败:', chrome.runtime.lastError?.message);
          return;
        }

        try {
          const body = result.body;
          if (!body.includes('access_token')) {
            console.log('[Ask Marcel] OAuth 响应中没有 access_token');
            return;
          }

          const data = JSON.parse(body);
          const token = data.access_token;
          if (!token || !token.startsWith('eyJ')) {
            console.log('[Ask Marcel] access_token 格式无效');
            return;
          }

          const payload = decodeJwtPayload(token);
          if (!payload) {
            console.log('[Ask Marcel] JWT 解析失败');
            return;
          }

          const appid = payload.appid;
          const aud = payload.aud;
          console.log('[Ask Marcel] OAuth token: appid=' + appid + ', aud=' + aud);

          // 1. Graph token（基本 Teams token）
          if (!capturedGraph && (aud === GRAPH_AUD || aud === GRAPH_AUD_ALT)) {
            capturedGraph = token;
            console.log('[Ask Marcel] ✓ 捕获到 Graph token');
            scheduleSend();
          }

          // 2. chatsvcagg token（Teams 聊天记录必需）
          if (!capturedChatsvcagg && appid === BASIC_TEAMS_APP_ID && aud === CHATSVCAGG_AUD) {
            capturedChatsvcagg = token;
            console.log('[Ask Marcel] ✓ 捕获到 chatsvcagg token');
            scheduleSend();
          }

          // 3. IC3 token（Teams 聊天历史必需）
          if (!capturedIc3 && appid === BASIC_TEAMS_APP_ID && aud === IC3_AUD) {
            capturedIc3 = token;
            console.log('[Ask Marcel] ✓ 捕获到 IC3 token');
            scheduleSend();
          }

          // 4. Elevated token（历史版本下载必需）
          if (!capturedElevated && ELEVATED_APP_IDS.includes(appid) && (aud === GRAPH_AUD || aud === GRAPH_AUD_ALT)) {
            capturedElevated = token;
            console.log('[Ask Marcel] ✓ 捕获到 elevated token');
            scheduleSend();
          }

        } catch (e) {
          console.log('[Ask Marcel] OAuth 响应解析失败:', e.message);
        }
      }
    );
  }
});

chrome.debugger.onDetach.addListener((source) => {
  if (source.tabId === debuggedTabId) {
    console.log('[Ask Marcel] Debugger detached');
    debuggedTabId = null;
  }
});

chrome.runtime.onInstalled.addListener(() => {
  console.log('[Ask Marcel] ✓ Extension v3.2 installed');
});
