// atelier Java asset: the success arm of Result (hard rule 16).
package com.example.app.domain;

public record Ok<T, E>(T value) implements Result<T, E> {}
