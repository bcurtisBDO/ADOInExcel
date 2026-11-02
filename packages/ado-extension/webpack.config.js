const CopyWebpackPlugin = require("copy-webpack-plugin");
const path = require("path");

const azureDevOpsSdk = require.resolve("azure-devops-extension-sdk");

module.exports = (_env, argv) => {
  const production = argv.mode === "production";
  return {
    devtool: production ? false : "inline-source-map",
    devServer: {
      server: "https",
      port: 3000,
      open: false,
      hot: true,
      static: {
        directory: path.join(__dirname, "./")
      }
    },
    entry: {
      index: "./src/index.ts",
      download: "./src/download.ts"
    },
    output: {
      publicPath: "auto",
      filename: "[name].js",
      path: path.resolve(__dirname, "dist")
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
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
        }
      ]
    },
    resolve: {
      extensions: [".tsx", ".ts", ".js"],
      // azure-devops-extension-api requires the CJS SDK; our imports resolve the ESM
      // build. Webpack would otherwise ship both and trip the SDK's single-load guard.
      alias: {
        "@ado-open-in-excel/shared": path.resolve(__dirname, "../shared/src"),
        "azure-devops-extension-sdk": azureDevOpsSdk,
        "azure-devops-extension-sdk$": azureDevOpsSdk
      },
      fallback: {
        fs: false,
        stream: false,
        crypto: false
      }
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [{ from: "**/*.html", context: "src" }]
      })
    ]
  };
};
