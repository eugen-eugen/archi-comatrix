const path = require("path");
const webpack = require("webpack");

module.exports = {
  mode: "production",
  entry: "./src/main/comatrix-bundled.js",
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "comatrix-bundled.ajs",
    library: {
      type: "var",
      name: "ComatrixLib",
    },
  },
  target: "web",
  optimization: {
    minimize: false,
  },
  resolve: {
    alias: {
      crypto: path.resolve(__dirname, "src/main/crypto-polyfill.js"),
    },
    fallback: {
      stream: require.resolve("stream-browserify"),
      buffer: require.resolve("buffer/"),
      util: require.resolve("util/"),
      path: require.resolve("path-browserify"),
      fs: false,
      events: require.resolve("events/"),
    },
    fullySpecified: false,
  },
  plugins: [
    new webpack.DefinePlugin({
      "process.env.NODE_ENV": JSON.stringify("production"),
      global: "globalThis",
    }),
  ],
};
