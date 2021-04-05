/**
 * Replaces the angular brackets with the respective codes.
 *
 * @param {string} node String to be replaced
 * @returns {string} String with replaced angular brackets
 */
const sanitizeHtmlChars = node => {
  return node.replace(/>/g, '&gt;').replace(/</g, '&lt;');
};

export default sanitizeHtmlChars;
