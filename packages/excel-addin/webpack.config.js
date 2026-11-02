const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  return {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts",
      auth: "./src/auth/authDialog.ts"
    },
    output: {
      clean: true,
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js"
    },
    resolve: {
      extensions: [".ts", ".js"],
      alias: {
        "@ado-open-in-excel/shared": path.resolve(__dirname, "../shared/src")
      }
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          include: [
            path.resolve(__dirname, "src"),
            path.resolve(__dirname, "../shared/src")
          ],
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env", "@babel/preset-typescript"]
            }
          }
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]"
          }
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"]
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new HtmlWebpackPlugin({
        filename: "auth.html",
        template: "./src/auth/auth.html",
        chunks: ["polyfill", "auth"]
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets",
            to: "assets"
          },
          {
            from: "src/taskpane/taskpane.css",
            to: "taskpane.css"
          },
          {
            from: "manifest.xml",
            to: "manifest.xml"
          }
        ]
      })
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions()
      },
      port: process.env.npm_package_config_dev_server_port || 3001
    }
  };
};
