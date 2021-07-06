function quoteToDoubleQuote(string) {
  return string.replace(/"/g, '""');
}

module.exports = {
  quoteToDoubleQuote
}
