// The prebuilt browser bundle avoids Node's Buffer, which mammoth's lib/ entry needs.
declare module 'mammoth/mammoth.browser.js' {
  import mammoth = require('mammoth');
  export = mammoth;
}
