const Dotenv = require('dotenv-webpack');
module.exports = require('@battis/gas-lighter/webpack')({
  root: __dirname,
  plugins: [new Dotenv()],
  production: true
});
