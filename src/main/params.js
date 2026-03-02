/**
 * params.js
 * Parameter handling utilities for jArchi scripts
 *
 * Parameters can be passed to scripts using the --paramName format:
 * Example: jArchi.script("comatrix.ajs", "--baselineModel", "/path/to/model.archimate")
 */

/**
 * Gets the value of a named parameter from command line arguments
 * @param {string} paramName - The parameter name (without -- prefix)
 * @returns {string|null} The parameter value if found, null otherwise
 *
 * @example
 * // If script is called with: --baselineModel /path/to/model.archimate
 * const path = getParameter("baselineModel");
 * // Returns: "/path/to/model.archimate"
 */
function getParameter(paramName) {
  const argv = $.process.argv;

  if (!argv || argv.length === 0) {
    return null;
  }

  const paramFlag = `--${paramName}`;

  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === paramFlag && i + 1 < argv.length) {
      return argv[i + 1];
    }
  }

  return null;
}

/**
 * Checks if a parameter exists in command line arguments
 * @param {string} paramName - The parameter name (without -- prefix)
 * @returns {boolean} True if parameter exists, false otherwise
 */
function hasParameter(paramName) {
  const argv = $.process.argv;

  if (!argv || argv.length === 0) {
    return false;
  }

  const paramFlag = `--${paramName}`;
  return argv.includes(paramFlag);
}

/**
 * Gets all parameters as an object
 * @returns {Object} Object with parameter names as keys and their values
 *
 * @example
 * // If script is called with: --baselineModel /path/to/model.archimate --verbose
 * const params = getAllParameters();
 * // Returns: { baselineModel: "/path/to/model.archimate", verbose: null }
 */
function getAllParameters() {
  const params = {};
  const argv = $.process.argv;

  if (!argv || argv.length === 0) {
    return params;
  }

  for (let i = 0; i < argv.length; i++) {
    if (argv[i].startsWith("--")) {
      const paramName = argv[i].substring(2);
      const paramValue = i + 1 < argv.length && !argv[i + 1].startsWith("--") ? argv[i + 1] : null;
      params[paramName] = paramValue;
      if (paramValue !== null) {
        i++; // Skip the value in next iteration
      }
    }
  }

  return params;
}

module.exports = {
  getParameter,
  hasParameter,
  getAllParameters,
};
