/**
 * Strips off the quotation marks from the first and last position.
 *
 * @param {*} value Value to sanitize.
 * @returns {string} Sanitized Value after replacing quotation marks
 */
const sanitizeStringValues = value => {
  if (typeof value === 'string') {
    const len = value.length;
    if (value[0]=='"' && value[len-1]=='"') {
      return value.substr(1, len-2);
    }
  }
  return value;
};

export default sanitizeStringValues;
