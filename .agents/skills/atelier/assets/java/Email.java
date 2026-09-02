// atelier Java asset: value-record exemplar for a branded primitive at a trust
// boundary (hard rule 12, Java expression). The compact constructor is the
// guard (constructing an invalid instance is a bug, so it throws); the static
// parse is the boundary factory returning Result (expected-invalid input is a
// value, not an exception). Copy this shape for Money (integer minor units),
// UserId, IsoCountryCode, and every other domain primitive.
package com.example.app.domain;

public record Email(String value) {
  private static final java.util.regex.Pattern SHAPE = java.util.regex.Pattern.compile("^[^@\\s]+@[^@\\s]+$");

  public Email {
    if (!SHAPE.matcher(value).matches()) {
      throw new IllegalArgumentException("email");
    }
  }

  public static Result<Email, String> parse(String raw) {
    return SHAPE.matcher(raw).matches() ? new Ok<>(new Email(raw)) : new Err<>("invalid_email");
  }
}
