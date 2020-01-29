/*
    <copyright file="router.tsx" company="Microsoft Corporation">
    Copyright (c) Microsoft Corporation. All rights reserved.
    </copyright>
*/

import React from 'react';
import { Route } from "react-router-dom";
import AddFavorites from "../components/add-favorites";
import OtherRoom from "../components/other-room";

const ReactRouter = () => {
    return (
        <React.Fragment>
            <Route path="/Meeting/OtherRoom/" component={OtherRoom} />
            <Route path="/Meeting/AddFavourite/" component={AddFavorites} />
        </React.Fragment>
    );
}
export default ReactRouter;
