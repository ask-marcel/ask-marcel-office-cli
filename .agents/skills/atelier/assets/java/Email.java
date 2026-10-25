// atelier Java asset: value-record exemplar for a branded primitive at a trust
// boundary (hard rule 12, Java expression). The compact constructor is the
// guard (constructing an invalid instance is a bug, so it throws); the static
// parse is the boundary factory returning Result (expected-invalid input is a
// value, not an exception), and its error is a nested enum, never a bare
// String, so a caller switches over the cases the type names (rule 16). Copy
// this shape for Money (integer minor units), UserId, IsoCountryCode, and
// every other domain primitive.
package com.example.app.domain;

import java.util.regex.Pattern;

public record Email(String value) {
  private static final Pattern SHAPE = Pattern.compile("^[^@\\s]+@[^@\\s]+$");

  public Email {
    if (!SHAPE.matcher(value).matches()) {
      throw new IllegalArgumentException("email");
    }
  }

  public enum Error {
    MALFORMED
  }

  public static Result<Email, Error> parse(String raw) {
    return SHAPE.matcher(raw).matches() ? new Ok<>(new Email(raw)) : new Err<>(Error.MALFORMED);
  }
}
