const path = require("path");
const webpack = require("webpack");

module.exports = {
  mode: "production",
  entry: {
    "comatrix-bundled": "./src/main/comatrix.js",
    "applist-bundled": "./src/main/applist.js",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].ajs",
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
