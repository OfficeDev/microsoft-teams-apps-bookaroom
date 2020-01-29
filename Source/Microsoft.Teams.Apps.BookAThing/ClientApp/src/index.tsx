/*
    <copyright file="index.tsx" company="Microsoft Corporation">
    Copyright (c) Microsoft Corporation. All rights reserved.
    </copyright>
*/

import React from 'react';
import ReactDOM from 'react-dom';
import { BrowserRouter as Router } from 'react-router-dom';
import ReactRouter from './router/router';

ReactDOM.render(
    <Router>
        <ReactRouter />
    </Router>, document.getElementById('root'));