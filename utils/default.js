/**
 * Handlize the title
 * @param {string} title
 * @returns {string}
 */
const handlize = (title, separator = "-") => {
  return `${title}`
    .toLowerCase()
    .replace(/ /g, separator)
    .replace(/[^a-z0-9-]/g, "");
};

/**
 * Create key from title
 * @param {string} title
 * @returns {string}
 */
const createKey = (title) => {
  return `${title}`.toLowerCase().replace(/[^a-z0-9]/g, "");
};

/**
 * Convert new lines to HTML breaks
 * @param {string} text
 * @param {boolean} xhtml
 * @returns {string}
 */
const nl2br = (text, xhtml = true) => {
  if (text === undefined || text === null) return "";
  const breakTag = xhtml ? "<br/>" : "<br>";
  return String(text).replace(/\r\n|\r|\n/g, breakTag + "\n");
};

module.exports = {
  handlize,
  nl2br,
  createKey,
};
