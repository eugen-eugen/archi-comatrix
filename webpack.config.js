const path = require("path");
const webpack = require("webpack");

module.exports = {
  mode: "production",
  entry: "./src/main/comatrix.js",
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
    fallback: {
      path: require.resolve("path-browserify"),
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
