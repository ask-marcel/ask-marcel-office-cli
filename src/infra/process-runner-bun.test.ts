import { describeProcessRunnerContract } from '../test-helpers/process-runner-contract.ts';
import { createBunProcessRunner } from './process-runner-bun.ts';

describeProcessRunnerContract('Bun process runner adapter', createBunProcessRunner);
