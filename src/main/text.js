/**
 * text.js
 * Small string helpers shared across the comatrix scripts.
 */

/**
 * Returns true when the value is a non-empty string once trimmed
 * (i.e. not null/undefined, and not blank/whitespace-only).
 * @param {*} value - The value to test
 * @returns {boolean}
 */
function hasText(value) {
  return value != null && String(value).trim() !== "";
}

/**
 * Returns true when the value is null/undefined or blank/whitespace-only.
 * @param {*} value - The value to test
 * @returns {boolean}
 */
function isBlank(value) {
  return !hasText(value);
}

module.exports = {
  hasText,
  isBlank,
};
