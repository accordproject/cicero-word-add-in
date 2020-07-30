const sanitizeHtmlChars = node => {
  return node.replace(/>/g, '&gt;').replace(/</g, '&lt;');
};

export default sanitizeHtmlChars;
