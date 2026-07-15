// atelier Java asset: the failure arm of Result (hard rule 16). A business
// failure is a value here, never a thrown exception (rule 10).
package com.example.app.domain;

public record Err<T, E>(E error) implements Result<T, E> {}
