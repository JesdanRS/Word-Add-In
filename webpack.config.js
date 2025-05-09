/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");
const path = require("path");
const DotenvWebpackPlugin = require("dotenv-webpack");

const urlDev = "https://localhost:3000/";
const urlProd = "https://mrbl4ck.pages.gitlab.io/word-add-in/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: ["./src/taskpane/index.tsx", "./src/taskpane/index.html"],
    },
    output: {
      clean: true,
      path: path.resolve(__dirname, "public"),
      publicPath: "/word-add-in",
    },
    optimization: {
      minimize: false,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          include: path.resolve(__dirname, "src"),
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-typescript"],
            },
          },
        },
        {
          test: /\.tsx?$/,
          include: path.resolve(__dirname, "src"),
          use: ["ts-loader"],
        },
        {
          test: /\.html$/,
          include: path.resolve(__dirname, "src"),
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          include: path.resolve(__dirname, "assets"),
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new DotenvWebpackPlugin(),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "index.html",
        template: "./src/taskpane/index.html",
        chunks: ["taskpane"],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      devMiddleware: {
        publicPath: "/word-add-in",
      },
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
