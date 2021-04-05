/**
 * Generates a title from the variable using the title and type.
 *
 * @param {string} title Title of the variable. E.g. Receiver-1, Shipper-1
 * @param {string} type  Type of the variable
 * @returns {string} New title combining title and type
 */
const titleGenerator = (title, type) => {
  return `${title} | ${type}`;
};

export default titleGenerator;
