// atelier Java asset: the sealed Result union (hard rule 16, Java expression).
// Copy into your domain package; rename `com.example.app.domain` to your own.
// Ok/Err complete the union; use-cases pattern-match with an exhaustive switch.
package com.example.app.domain;

public sealed interface Result<T, E> permits Ok, Err {}
