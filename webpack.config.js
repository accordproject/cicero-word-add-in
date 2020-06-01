const devCerts = require('office-addin-dev-certs');
const { CleanWebpackPlugin } = require('clean-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require('webpack');

module.exports = async (env, options)  => {
  const dev = options.mode === 'development';
  const config = {
    devtool: 'source-map',
    entry: {
      polyfill: 'babel-polyfill',
      main: './src/index.js',
    },
    resolve: {
      extensions: ['.html', '.js'],
    },
    module: {
      rules: [
        {
          test: /\.jsx?$/,
          use: [
              'react-hot-loader/webpack',
              'babel-loader',
          ],
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader'],
        },
        {
          test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
          use: {
              loader: 'file-loader',
              query: {
                  name: 'assets/[name].[ext]',
                },
              },  
            },   
          ],
    },    
    plugins: [
      new CleanWebpackPlugin(),
      new CopyWebpackPlugin({
        patterns: [
          { to: 'index.css', from: './src/index.css' },
        ],
      }),
      new ExtractTextPlugin('[name].[hash].css'),
      new HtmlWebpackPlugin({
        filename: 'index.html',
        template: './src/index.html',
      }),
      new CopyWebpackPlugin({
        patterns: [
          { to: './assets', from: 'assets' },
        ],
      }),
    ],
    devServer: {
      headers: {
        'Access-Control-Allow-Origin': '*',
      },      
      https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
