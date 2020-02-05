/*
    <copyright file="add-favorites.tsx" company="Microsoft Corporation">
    Copyright (c) Microsoft Corporation. All rights reserved.
    </copyright>
*/

import * as React from "react";
import "./theme.css";
import { Button, Loader, Divider, List, Icon, Text, Provider, themes, Dropdown, Flex } from '@fluentui/react';
import { components } from "react-select";
import AsyncSelect from "react-select/async";
import * as microsoftTeams from "@microsoft/teams-js";
import * as Constants from "../constants";
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";
let moment = require("moment");
import * as timezone from 'moment-timezone';

const browserHistory = createBrowserHistory({ basename: '' });
interface IAddFavoriteProps { }
let reactPlugin = new ReactPlugin();

/** User favorite room */
class FavoriteRoom {
    UserAdObjectId?: string | null = null;
    RoomEmail?: string | null = null;
    RoomName?: string | null = null;
    BuildingName?: string | null = null;
    BuildingEmail?: string | null = null;
}

/** State interface. */
interface IState {
    /**Favorite rooms list. */
    favoriteRooms: Array<FavoriteRoom>,
    /**Selected room object. */
    selectedRoom: any,
    /**Loading icon visibility. */
    loading: boolean,
    /**Add button disable/enable. */
    addDisable: boolean,
    /**Error message text. */
    message?: string | null,
    /**Error message visibility. */
    showMessage: boolean,
    /**Error message color. */
    messageColor?: string,
    /**Selected theme. */
    theme: any,
    /**Is user authorized. */
    authorized: boolean,
    /**Loading for favorite list. */
    loadingFavoriteList: boolean,
    /**Top five rooms to display in dropdown */
    topFiveRooms: Array<any>,
    /**Supported time zones for user */
    supportedTimeZones: Array<any>,
    /**Selected time zone for user */
    selectedTimeZone: any,
    /**Boolean indicating if time zones are loading in dropdown */
    timeZonesLoading: boolean,
    resourceStrings: any,
    resourceStringsLoaded: boolean,
    isRoomDeleted: boolean,
    errorResponseDetail: IErrorResponse,
};

/**Server error response interface */
interface IErrorResponse {
    statusCode?: string,
    errorMessage?: string,
}

/** Component for managing user favorites. */
class AddFavorites extends React.Component<IAddFavoriteProps, IState>
{
    /**Reply to activity Id. */
    replyTo?: string | null = null;
    /**Auth token. */
    token?: string | null = null;
    /**Component state. */
    state: IState;
    /** Theme color according to teams theme*/
    themeColor?: any = undefined;
    /** Theme styles according to teams theme*/
    themeStyle?: any = undefined;
    /** Instrumentation key for telemetry logging*/
    telemetry: any = undefined;
    appInsights: ApplicationInsights;
    userObjectId: any;
    userTimeZone: any = null;
    strings: any = {};

    /**
     * Contructor to initialize component.
     * @param props Props of component.
     */
    constructor(props: IAddFavoriteProps) {
        super(props);
        this.state = {
            favoriteRooms: [],
            selectedRoom: null,
            loading: false,
            addDisable: true,
            message: null,
            showMessage: false,
            messageColor: undefined,
            theme: null,
            authorized: true,
            loadingFavoriteList: false,
            topFiveRooms: [],
            supportedTimeZones: [],
            selectedTimeZone: null,
            timeZonesLoading: false,
            resourceStrings: {},
            resourceStringsLoaded: false,
            isRoomDeleted: false,
            errorResponseDetail: {
                errorMessage: undefined,
                statusCode: undefined,
            },
        };

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.replyTo = params.get("replyTo");
        this.telemetry = params.get("telemetry");
        this.token = params.get("token");
        this.strings = JSON.parse(localStorage.getItem("Strings")!);
        this.appInsights = new ApplicationInsights({
            config: {
                instrumentationKey: this.telemetry,
                extensions: [reactPlugin],
                extensionConfig: {
                    [reactPlugin.identifier]: { history: browserHistory }
                }
            }
        });
        this.appInsights.loadAppInsights();
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        this.setState({ loading: true });

        // Call the initialize API first
        microsoftTeams.initialize();

        // Check the initial theme user chose and respect it
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            if (context && context.theme) {
                this.setState({ theme: context.theme });
                if (context.theme === Constants.DarkTheme) {
                    this.themeColor = Constants.DarkThemeColors;
                    this.themeStyle = Constants.DarkStyles;
                }
                else if (context.theme === Constants.DefaultTheme) {
                    this.themeColor = Constants.DefaultThemeColors;
                    this.themeStyle = Constants.DefaultStyles;
                }
                else {
                    this.themeColor = Constants.ContrastThemeColors;
                    this.themeStyle = Constants.ContrastStyles;
                }
            }
        });

        // Handle theme changes
        microsoftTeams.registerOnThemeChangeHandler((theme) => {
            this.setState({ theme: theme });
            if (theme === Constants.DarkTheme) {
                this.themeColor = Constants.DarkThemeColors;
                this.themeStyle = Constants.DarkStyles;
            }
            else if (theme === Constants.DefaultTheme) {
                this.themeColor = Constants.DefaultThemeColors;
                this.themeStyle = Constants.DefaultStyles;
            }
            else {
                this.themeColor = Constants.ContrastThemeColors;
                this.themeStyle = Constants.ContrastStyles;
            }
        });
        this.getResourceStrings();
        this.userTimeZone = timezone.tz.guess(true);
        this.getSavedTimeZone();
        this.getTopNRooms();
        this.getFavoriteRooms();
    }

    /** 
    *  Fetch resource strings.
    * */
    getResourceStrings = async () => {
        this.setState({ loading: true });
        let request = new Request("/api/ResourcesApi/GetResourceStrings", {
            headers: new Headers({
                "Authorization": "Bearer " + this.token
            })
        });
        const resourceStrings = await fetch(request);
        this.setState({ loading: false });

        if (resourceStrings.status === 401) {
            this.setState({ authorized: false, loading: false });
            this.appInsights.trackEvent({ name: `Unauthorized` }, { User: this.userObjectId });
        }
        else if (resourceStrings.status === 200) {
            const response = await resourceStrings.json();
            if (response !== null) {
                this.setState({ resourceStrings: response, resourceStringsLoaded: true });
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'GetResourceStringsAsync' - Request failed:${resourceStrings.status}`, severityLevel: SeverityLevel.Warning });
            this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
        }
    }

    /** 
     *  Fetch supported time zones for user from Exchange.
     * */
    getSupportedTimeZones = async () => {
        this.setState({ timeZonesLoading: true });
        let request = new Request("/api/MeetingApi/GetSupportedTimeZonesAsync", {
            headers: new Headers({
                "Authorization": "Bearer " + this.token
            })
        });

        const supportedTimezoneResponse = await fetch(request);
        this.setState({ timeZonesLoading: false });

        if (supportedTimezoneResponse.status === 401) {
            const response = await supportedTimezoneResponse.json();
            if (response) {
                this.setState({
                    errorResponseDetail: {
                        errorMessage: response.message,
                        statusCode: response.code,
                    }
                })
            }

            this.setState({ authorized: false });
            this.appInsights.trackEvent({ name: `Unauthorized` }, { User: this.userObjectId });
        }
        else if (supportedTimezoneResponse.status === 200) {
            const response = await supportedTimezoneResponse.json();
            if (response !== null) {
                this.setState({ supportedTimeZones: response });
                let self = this;
                let tzResult = self.state.supportedTimeZones.find(function (tz) { return tz === self.userTimeZone });
                if (tzResult) {
                    this.setState({ selectedTimeZone: self.userTimeZone });
                    this.saveUserTimeZone(self.userTimeZone);
                }
                else {
                    this.setMessage(this.state.resourceStrings.TimezoneNotSupported, Constants.ErrorMessageRedColor, false);
                }
            }
            else {
                this.appInsights.trackTrace({ message: `'GetSupportedTimezones' - Request failed:${supportedTimezoneResponse.status}`, severityLevel: SeverityLevel.Warning });
                this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'TopFiveRoomsAsync' - Request failed:${supportedTimezoneResponse.status}`, severityLevel: SeverityLevel.Warning });
            this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
        }
    }

    /** 
     *  Fetch saved timezone of user.
     * */
    getSavedTimeZone = async () => {
        let request = new Request("/api/MeetingApi/GetUserTimeZoneAsync", {
            headers: new Headers({
                "Authorization": "Bearer " + this.token
            })
        });

        const timezoneResponse = await fetch(request);
        if (timezoneResponse.status === 401) {
            this.setState({ authorized: false });
            this.appInsights.trackEvent({ name: `Unauthorized` }, { User: this.userObjectId });
        }
        else if (timezoneResponse.status === 200) {
            const response = await timezoneResponse.json();
            if (response) {
                this.setState({ selectedTimeZone: response.IanaTimezone });
                this.userTimeZone = response.IanaTimezone;
            }
            else {
                this.getSupportedTimeZones();
            }
        }
        else if (timezoneResponse.status === 204) {
            this.getSupportedTimeZones();
        }
        else {
            this.appInsights.trackTrace({ message: `'TopFiveRoomsAsync' - Request failed:${timezoneResponse.status}`, severityLevel: SeverityLevel.Warning });
            this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
        }
    }

    /** 
    *  Save timezone selected by user.
    *  @param timezone  Selected timezone name.
    * */
    saveUserTimeZone = async (timezone: string) => {
        this.appInsights.trackTrace({ message: `User ${this.userObjectId} saving timezone ${timezone}` });
        const saveTimeZoneResult = await fetch("/api/MeetingApi/SaveTimeZoneAsync", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + this.token
            },
            body: JSON.stringify({ PartitionKey: "msteams", IanaTimezone: timezone })
        });

        if (saveTimeZoneResult.status === 401) {
            this.appInsights.trackEvent({ name: `Unauthorized` }, { User: this.userObjectId });
            this.setState({ authorized: false });
        }
        else if (saveTimeZoneResult.status !== 200) {
            this.appInsights.trackEvent({ name: `Updated user timezone` }, { User: this.userObjectId, Timezone: timezone });
            this.appInsights.trackTrace({ message: `'SaveUserTimezone' - Request failed:${saveTimeZoneResult.status}`, severityLevel: SeverityLevel.Warning });
            this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
        }
    }

    /**Get list of top N rooms to display in dropdown on click. */
    getTopNRooms = async () => {
        let dateTime = moment().utc().format("YYYY-MM-DD HH:mm:ss");
        let rooms = { Query: "", Duration: 0, TimeZone: this.state.selectedTimeZone, Time: dateTime, ScheduleRequired: false };
        const res = await fetch("/api/MeetingApi/TopNRoomsAsync", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + this.token
            },
            body: JSON.stringify(rooms)
        });

        if (res.status === 401) {
            const response = await res.json();
            if (response) {
                this.setState({
                    errorResponseDetail: {
                        errorMessage: response.message,
                        statusCode: response.code,
                    }
                })
            }

            this.appInsights.trackEvent({ name: `Unauthorized` }, { User: this.userObjectId });
            this.setState({ authorized: false });
        }
        else if (res.status === 200) {
            let response = await res.json();
            this.setState({ topFiveRooms: response });
        }
        else {
            this.appInsights.trackTrace({ message: `'TopFiveRoomsAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
            this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
        }
    }

    /** Render rooms dropdown. */
    renderRoomsDropdown() {
        const Option = (props: any) => {
            return (
                <components.Option {...props}>
                    <div className="addfav-dropdownlabel"><b className="addfav-dropdownlabel-b">{props.data.label}</b><b className="addfav-dropdownlabel-b2" style={{ color: props.data.Status === "Available" ? "green" : "red" }}>{props.data.Status}</b></div><br />
                    <div className="addfav-dropdownsublabel">{props.data.sublabel}</div>
                </components.Option>
            );
        };

        return (
            <div style={{ width: "inherit" }}>
                <AsyncSelect
                    defaultOptions={this.state.topFiveRooms}
                    placeholder={this.state.resourceStrings.SearchRoomDropdownPlaceholder}
                    components={{ Option, IndicatorSeparator: () => null }}
                    styles={this.themeStyle}
                    loadOptions={this.promiseOptions}
                    value={this.state.selectedRoom}
                    onChange={this.handleRoomChange}
                    theme={this.themeColor}
                />
            </div>
        )
    }

    /** Render favorite rooms list. */
    renderFavoriteList() {
        let self = this;
        if (self.state.loadingFavoriteList) {
            return (
                <Flex gap="gap.large">
                    <Flex
                        hAlign="center"
                        vAlign="center"
                        style={{
                            width: '90vw',
                            height: '50vh',
                        }}
                    >
                        <Loader />
                    </Flex>
                </Flex>
            );
        }
        else {
            if (this.state.favoriteRooms !== null && this.state.favoriteRooms.length > 0) {
                return (
                    <List items={this.state.favoriteRooms.map((room: any, index) => {
                        return (
                            {
                                key: index.toString(),
                                endMedia: <div id={'div' + index.toString()} onClick={() => this.removeRoom(index)}><Icon name="star" color="yellow" size="large" /></div>,
                                header: room.RoomName,
                                content: room.BuildingName,
                                styles: { padding: 0 }
                            }
                        )
                    })} />
                );
            }
            else {
                return (
                    <Flex gap="gap.small">
                        <Flex.Item>
                            <div
                                style={{
                                    position: 'relative',
                                }}
                            >
                                <Icon outline color="green" name="question-circle" />
                            </div>
                        </Flex.Item>

                        <Flex.Item grow>
                            <Flex column gap="gap.small" vAlign="stretch">
                                <div>
                                    <Text weight="bold" content={this.state.resourceStrings.NoFavoriteRoomsTaskModule} /><br />
                                    <Text content={this.state.resourceStrings.NoFavoritesDescriptionTaskModule} />
                                </div>
                            </Flex>
                        </Flex.Item>
                    </Flex>
                );

            }
        }
    }

    /** Get favorite rooms from table storage. */
    async getFavoriteRooms() {
        this.setState({ loadingFavoriteList: true });
        let request = new Request("/api/MeetingApi/GetFavoriteRoomsAsync", {
            headers: new Headers({
                "Authorization": "Bearer " + this.token
            })
        });
        const favoriteRooms = await fetch(request);

        if (favoriteRooms.status === 401) {
            const response = await favoriteRooms.json();
            if (response) {
                this.setState({
                    errorResponseDetail: {
                        errorMessage: response.message,
                        statusCode: response.code,
                    }
                })
            }

            this.appInsights.trackEvent({ name: `Unauthorized` }, { User: this.userObjectId });
            this.setState({ authorized: false, loadingFavoriteList: false });
        }
        else if (favoriteRooms.status === 200) {
            const response = await favoriteRooms.json();
            if (response !== null) {
                this.setState({ loadingFavoriteList: false, favoriteRooms: response });
            }
            else {
                this.setState({ loadingFavoriteList: false });
                this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'TopFiveRoomsAsync' - Request failed:${favoriteRooms.status}`, severityLevel: SeverityLevel.Warning });
            this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
        }
    }

    /** Submit selection and save changed to storage. */
    async submit() {
        this.appInsights.trackTrace({ message: `User ${this.userObjectId} submitted favorite list` });
        let self = this;
        self.setState({ loading: true });
        const favoriteSaveResult = await fetch("/api/MeetingApi/SubmitFavoritesAsync", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + this.token
            },
            body: JSON.stringify(self.state.favoriteRooms)
        });

        if (favoriteSaveResult.status === 401) {
            this.appInsights.trackEvent({ name: `Unauthorized` }, { User: this.userObjectId });
            this.setState({ authorized: false, loading: false });
        }
        else if (favoriteSaveResult.status === 200) {
            this.appInsights.trackEvent({ name: `Favorites updated` }, { User: this.userObjectId });
            if (favoriteSaveResult !== null) {
                let toBot = { Text: "fav closed", ReplyTo: this.replyTo };
                microsoftTeams.tasks.submitTask(toBot);
            }
            else {
                self.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'TopFiveRoomsAsync' - Request failed:${favoriteSaveResult.status}`, severityLevel: SeverityLevel.Warning });
            this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
        }
    }

    /**
     * Remove room from favorite.
     * @param index Index of room in array.
     */
    removeRoom = async (index: any) => {
        let favoritesList = this.state.favoriteRooms;
        let room = favoritesList[index];
        favoritesList.splice(index, 1);
        this.appInsights.trackEvent({ name: `Removed from favorite` }, { User: this.userObjectId, Room: room.RoomEmail });
        this.appInsights.trackTrace({ message: "User " + this.userObjectId + " removed room " + room.RoomName + " from favorites in client app" });
        this.setState({ favoriteRooms: favoritesList, showMessage: false, isRoomDeleted: true });
    }

    /** 
     *  Show message to user.
     *  @param message  Message to show.
     *  @param color Color of message text.
     *  @param loading  Disable or enable loading icon.
     * */
    setMessage = (message: string, color: string, loading: boolean) => {
        this.setState({ showMessage: true, message: message, loading: loading, messageColor: color });
    }

    /** Add room to favorite. */
    addRoom = async () => {
        let self = this;
        let selectedRoom = self.state.selectedRoom;
        if (selectedRoom) {
            let existing = self.state.favoriteRooms.find(function (room: any) {
                return room.RowKey === selectedRoom.value;
            });

            if (!existing) {
                if (self.state.favoriteRooms.length !== 15) {
                    let room = { PartitionKey: "", RowKey: selectedRoom.RowKey, RoomName: selectedRoom.RoomName, BuildingName: selectedRoom.BuildingName, BuildingEmail: selectedRoom.PartitionKey };
                    let favoritesList = this.state.favoriteRooms;
                    favoritesList.push(room);
                    this.appInsights.trackEvent({ name: `Added to favorite` }, { User: this.userObjectId, Room: selectedRoom.RowKey });
                    this.appInsights.trackTrace({ message: `User ${this.userObjectId} added room ${room.RoomName} to favorites in client app` });
                    this.setState({ favoriteRooms: favoritesList, selectedRoom: null });
                }
                else {
                    this.appInsights.trackTrace({ message: `User ${this.userObjectId} exceeded favorite rooms max count` });
                    this.setMessage(self.state.resourceStrings.CantAddMoreRooms, Constants.ErrorMessageRedColor, false);
                }
            }
            else {
                this.setMessage(self.state.resourceStrings.FavoriteRoomExist, Constants.ErrorMessageRedColor, false);
            }
        }
        else {
            this.setMessage(self.state.resourceStrings.SelectRoomToAdd, Constants.ErrorMessageRedColor, false);
        }

    }

    /**
     * Filter rooms according search input.
     * @param inputValue Input string.
     */
    filterRooms = async (inputValue: string) => {
        let dateTime = moment().utc().format("YYYY-MM-DD HH:mm:ss");
        if (inputValue) {
            let rooms = { Query: inputValue, Duration: 0, TimeZone: this.state.selectedTimeZone, Time: dateTime, ScheduleRequired: false };
            const searchedRooms = await fetch("/api/MeetingApi/SearchRoomAsync", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + this.token
                },
                body: JSON.stringify(rooms)
            });

            if (searchedRooms.status === 401) {
                const response = await searchedRooms.json();
                if (response) {
                    this.setState({
                        errorResponseDetail: {
                            errorMessage: response.message,
                            statusCode: response.code,
                        }
                    })
                }

                this.appInsights.trackEvent({ name: `Unauthorized` }, { User: this.userObjectId });
                this.setState({ authorized: false });
                return [];
            }
            else if (searchedRooms.status === 200) {
                return await searchedRooms.json();
            }
            else {
                this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                this.appInsights.trackTrace({ message: `'SearchRoomAsync' - Request failed:${searchedRooms.status}`, severityLevel: SeverityLevel.Warning });
            }
        }
        else {
            return [];
        }

    }

    /**
     * Handles asynchronous operation for searching room.
     * @param inputValue Input string.
     */
    promiseOptions = (inputValue: string) =>
        new Promise(async resolve => {
            resolve(await this.filterRooms(inputValue));
        });

    /**
     * Event called after selecting room.
     * @param optionSelected Selected room.
     */
    handleRoomChange = (optionSelected: any) => {
        this.setState({ selectedRoom: optionSelected, addDisable: false, showMessage: false });
    }

    /** Render validation message. */
    renderMessage() {
        if (this.state.showMessage === true) {
            return (
                <Text error content={this.state.message} />
            );
        }
        else {
            return (<Text error content="" />);
        }
    }

    /**
     * Event called after user selects timezone.
     * @param event Dropdown sythetic event object.
     * @param data Data props for dropdown.
     */
    handleTimezonSelectionChange = (event: React.SyntheticEvent<HTMLElement>, data?: any) => {
        let tzResult = this.state.supportedTimeZones.find(function (tz) { return tz === data.value });
        if (tzResult) {
            this.setState({ selectedTimeZone: data.value, showMessage: false });
            this.saveUserTimeZone(data.value);
        }
        else {
            this.setMessage(this.state.resourceStrings.TimezoneNotSupported, Constants.ErrorMessageRedColor, false);
        }
    }

    /**
     * Event called after user opens timezone dropdown.
     * @param e Dropdown sythetic event object.
     * @param {open}' Boolean indicating if dropdown is opened or closed.
     */
    handleOpenChange = (e: any, { open }: any) => {
        if (open && this.state.supportedTimeZones.length == 0) {
            this.getSupportedTimeZones();
        }
    }

    /** render unauthorized error messages based on status code */
    renderErrorMessage = () => {
        if (this.state.errorResponseDetail.statusCode === "signinRequired") {
            return (
                <Text content={this.state.resourceStrings.SignInErrorMessage} style={{ color: Constants.ErrorMessageRedColor }} />
            );
        }

        return <Text content={this.state.resourceStrings.InvalidTenant} style={{ color: Constants.ErrorMessageRedColor }} />
    }

    /** Render function. */
    render() {
        let self = this;
        const checkAuthAndRender = function () {
            if (self.state.authorized) {
                if (self.state.resourceStringsLoaded) {
                    return (
                        <Provider theme={self.state.theme === Constants.DefaultTheme ? themes.teams : self.state.theme === Constants.DarkTheme ? themes.teamsDark : themes.teamsHighContrast}>
                            <div className="containerdiv">
                                <Flex gap="gap.small" vAlign="center" styles={{ paddingBottom: '1rem' }}>
                                    <Flex.Item push>
                                        <Text content={self.state.resourceStrings.Timezone} />
                                    </Flex.Item>
                                    <Flex.Item size="size.quarter">
                                        <Dropdown
                                            items={self.state.supportedTimeZones}
                                            placeholder={self.state.resourceStrings.SelectTimezone}
                                            loading={self.state.timeZonesLoading}
                                            loadingMessage={self.state.resourceStrings.LoadingMessage}
                                            onOpenChange={self.handleOpenChange}
                                            fluid={true}
                                            onSelectedChange={self.handleTimezonSelectionChange}
                                            value={self.state.selectedTimeZone}
                                        />
                                    </Flex.Item>
                                </Flex>
                                <Flex gap="gap.small">
                                    <Text weight="bold" content={self.state.resourceStrings.Location} />
                                </Flex>
                                <Flex gap="gap.small">
                                    <Flex.Item grow>
                                        {self.renderRoomsDropdown()}
                                    </Flex.Item>

                                    <Button primary disabled={self.state.addDisable} onClick={() => self.addRoom()}>{self.state.resourceStrings.AddButton}</Button>
                                </Flex>
                                <Divider style={{ marginTop: '1rem', marginBottom: '1rem' }} />
                                <Flex gap="gap.small">
                                    <Flex.Item grow>
                                        <Flex>
                                            <Flex.Item grow>
                                                <div className="container-subdiv">
                                                    {self.renderFavoriteList()}
                                                </div>
                                            </Flex.Item>
                                        </Flex>
                                    </Flex.Item>
                                </Flex>
                                <div className="footer">
                                    <Flex gap="gap.small">
                                        <Flex.Item grow>
                                            {self.renderMessage()}
                                        </Flex.Item>
                                        <Button loading={self.state.loading} disabled={self.state.selectedTimeZone === null || self.state.loading === true || self.state.favoriteRooms === null || (self.state.favoriteRooms.length === 0 && self.state.isRoomDeleted === false)} primary onClick={() => self.submit()} content={self.state.resourceStrings.DoneButton} />
                                    </Flex>
                                </div>
                            </div>
                        </Provider>
                    );
                }
                else {
                    return (<Loader />);
                }
            }
            else {
                return (
                    <Provider theme={self.state.theme === Constants.DefaultTheme ? themes.teams : self.state.theme === Constants.DarkTheme ? themes.teamsDark : themes.teamsHighContrast}>
                        <div className="containerdiv">
                            <div className="containerdiv-unauthorized">
                                <Flex gap="gap.small" vAlign="center" hAlign="center">
                                    {self.renderErrorMessage()}
                                </Flex>
                            </div>
                        </div>
                    </Provider>
                );
            }
        }

        return (checkAuthAndRender());
    }
}

export default withAITracking(reactPlugin, AddFavorites);