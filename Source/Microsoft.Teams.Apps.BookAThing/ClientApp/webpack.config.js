// <copyright file="webpack.config.js" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

const path = require("path");

module.exports = {
    mode: 'production',
    entry: {
        index: "./src/index.tsx"
    },
    //devtool: 'inline-source-map',
    devtool: false,
    resolve: { extensions: ['.js', '.jsx', '.ts', '.tsx'] },
    output: {
        path: path.resolve(__dirname, "../wwwroot/dist"),
        filename: "bundle.js"
    },
    module: {
        rules: [
            {
                use: {
                    loader: "babel-loader"
                },
                test: /\.js$/,
                exclude: /node_modules/ //excludes node_modules folder from being transpiled by babel. We do this because it's a waste of resources to do so.
            },
            {
                test: /\.(sa|sc|c)ss$/,
                use: ["style-loader", "css-loader"]
            },
            { test: /\.tsx?$/, include: /ClientApp/, use: 'awesome-typescript-loader?silent=true' },
            { test: /\.(png|jpg|jpeg|gif|svg)$/, use: 'url-loader?limit=25000' },
            { test: /\.(png|woff|woff2|eot|ttf)$/, use: 'url-loader?limit=100000' }

        ]
    }
};