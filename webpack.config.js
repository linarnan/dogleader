const path = require('path');

const HtmlWebpackPlugin = require('html-webpack-plugin');
const InlineSourceWebpackPlugin = require('inline-source-webpack-plugin');

module.exports = {
	mode: "development",
	entry: './src/main.js',
	output: {
		path: path.resolve(__dirname, './dist'),
		filename: 'index.html',
	},
	plugins: [
		new HtmlWebpackPlugin({
			template: 'src/index.html',
		}),
		new InlineSourceWebpackPlugin({
			compress: true,
			rootpath: './',
			noAssetMatch: 'warn'
		})
	]
};