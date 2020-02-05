/*
    <copyright file="other-room.tsx" company="Microsoft Corporation">
    Copyright (c) Microsoft Corporation. All rights reserved.
    </copyright>
*/

import * as React from "react";
import "./theme.css";
import { Button, Text, Provider, themes, Dropdown, Flex } from '@fluentui/react';
import Select, { components } from "react-select";
import AsyncSelect from "react-select/async";
import * as microsoftTeams from "@microsoft/teams-js";
import * as Constants from "../constants";
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";
let moment = require("moment");
import * as timezone from 'moment-timezone';
const browserHistory = createBrowserHistory({ basename: '' });
interface IOtherRoomProps { }
let reactPlugin = new ReactPlugin();

/** 
  *  State interface.
* */
interface IState {
    /**Duration list. */
    duration: Array<any>,
    /**Selected room object from dropdown. */
    selectedRoom: any,
    /**Selected duration object from dropdown. */
    selectedDuration: any
    /**Loading icon visibility. */
    loading: boolean,
    /**Search query. */
    searchQuery?: string | undefined,
    /**Error message text. */
    message?: string | null,
    /**Error message visibility. */
    showMessage: boolean,
    /**Error message color. */
    messageColor?: string
    /**Selected theme. */
    theme: any,
    /**Is user authorized. */
    authorized: boolean,
    /**Top five rooms to display in dropdown */
    topFiveRooms: Array<any>,
    /**Supported time zones for user */
    supportedTimeZones: Array<any>,
    /**Selected time zone for user */
    selectedTimeZone: any,
    /**Boolean indicating if time zones are loading in dropdown */
    timeZonesLoading: boolean,
    resourceStrings: any,
    errorResponseDetail: IErrorResponse,
};

/**Server error response interface */
interface IErrorResponse {
    statusCode?: string,
    errorMessage?: string,
}

/** 
 *  OtherRoom component.
**/
class OtherRoom extends React.Component<IOtherRoomProps, IState>
{
    /**Instrumentation key. */
    telemetry?: any = null;
    /**Reply to activirt Id. */
    replyTo?: string | null = null;
    /**Auth token. */
    token?: string | null = null;
    /**Component state. */
    state: IState;
    /** Theme color according to teams theme*/
    themeColor?: any = undefined;
    /** Theme styles according to teams theme*/
    themeStyle?: any = undefined;
    appInsights: ApplicationInsights;
    userObjectIdentifier: any;
    userTimeZone: any = null;

    /**
     * Constructor to initialize component.
     * @param props Props for component.
     */
    constructor(props: IOtherRoomProps) {
        super(props);
        this.state = {
            duration: Constants.DurationArray,
            selectedRoom: null,
            selectedDuration: null,
            loading: false,
            searchQuery: undefined,
            message: null,
            showMessage: false,
            messageColor: undefined,
            theme: null,
            authorized: true,
            topFiveRooms: [],
            supportedTimeZones: [],
            selectedTimeZone: null,
            timeZonesLoading: false,
            resourceStrings: {},
            errorResponseDetail: {
                errorMessage: undefined,
                statusCode: undefined,
            },
        };

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.replyTo = params.get("replyTo");
        this.token = params.get("token");

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
        // Call the initialize API first
        microsoftTeams.initialize();

        // Check the initial theme user chose and respect it
        microsoftTeams.getContext((context) => {
            this.userObjectIdentifier = context.userObjectId;

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
        this.setState({ selectedDuration: Constants.DurationArray[0] });
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
            this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} is unauthorized!`, severityLevel: SeverityLevel.Warning });
        }
        else if (resourceStrings.status === 200) {
            const response = await resourceStrings.json();
            if (response !== null) {
                this.setState({ resourceStrings: response });
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'GetResourceStringsAsync' - Request failed:${resourceStrings.status}`, severityLevel: SeverityLevel.Warning });
            this.setMessage("Something went wrong and I can’t do that right now. Try again in a few minutes.", Constants.ErrorMessageRedColor, false);
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

        const supportedTimeZones = await fetch(request);
        this.setState({ timeZonesLoading: false });

        if (supportedTimeZones.status === 401) {
            const response = await supportedTimeZones.json();
            if (response) {
                this.setState({
                    errorResponseDetail: {
                        errorMessage: response.message,
                        statusCode: response.code,
                    }
                })
            }

            this.setState({ authorized: false, loading: false });
            this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} is unauthorized!`, severityLevel: SeverityLevel.Warning });
        }
        else if (supportedTimeZones.status === 200) {
            const response = await supportedTimeZones.json();

            if (response !== null) {
                this.setState({ supportedTimeZones: response });
                let self = this;
                let tzResult = self.state.supportedTimeZones.find(function (tz) { return tz === self.userTimeZone });

                if (tzResult) {
                    this.setState({ selectedTimeZone: self.userTimeZone });
                    this.saveUserTimeZone(self.userTimeZone);
                    this.getTopNRooms();
                }
                else {
                    this.setMessage(this.state.resourceStrings.TimezoneNotSupported, Constants.ErrorMessageRedColor, false);
                }
            }
            else {
                this.appInsights.trackTrace({ message: `'GetSupportedTimezones' - Request failed:${supportedTimeZones.status}`, severityLevel: SeverityLevel.Warning });
                this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'TopFiveRoomsAsync' - Request failed:${supportedTimeZones.status}`, severityLevel: SeverityLevel.Warning });
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

        const savedTimeZone = await fetch(request);
        if (savedTimeZone.status === 401) {
            this.setState({ authorized: false, loading: false });
            this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} is unauthorized!`, severityLevel: SeverityLevel.Warning });
        }
        else if (savedTimeZone.status === 200) {
            const response = await savedTimeZone.json();
            if (response !== null) {
                this.setState({ selectedTimeZone: response.IanaTimezone });
                this.userTimeZone = response.IanaTimezone;
                this.getTopNRooms();
            }
            else {
                this.getSupportedTimeZones();
            }
        }
        else if (savedTimeZone.status === 204) {
            this.getSupportedTimeZones();
        }
        else {
            this.appInsights.trackTrace({ message: `'GetUserTimeZoneAsync' - Request failed:${savedTimeZone.status}`, severityLevel: SeverityLevel.Warning });
            this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
        }
    }

    /** 
    *  Save timezone selected by user.
    *  @param timezone  Selected timezone name.
    * */
    saveUserTimeZone = async (timezone: string) => {
        this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} saving timezone ${timezone}` });
        const timeZoneSaveResult = await fetch("/api/MeetingApi/SaveTimeZoneAsync", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + this.token
            },
            body: JSON.stringify({ PartitionKey: "msteams", IanaTimezone: timezone })
        });

        if (timeZoneSaveResult.status === 401) {
            this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} is unauthorized!`, severityLevel: SeverityLevel.Warning });
            this.setState({ authorized: false, loading: false });
        }
        else if (timeZoneSaveResult.status !== 200) {
            this.appInsights.trackTrace({ message: `'SaveUserTimezone' - Request failed:${timeZoneSaveResult.status}`, severityLevel: SeverityLevel.Warning });
            this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
        }
    }

    /** 
     *  Show message to user.
     *  @param message  Message to show.
     *  @param loading  Disable or enable loading icon.
     * */
    setMessage = (message: string, color: string, loading: boolean) => {
        this.setState({ showMessage: true, message: message, loading: loading, messageColor: color });
    }

    /**Get list of top N rooms to display in dropdown on click. */
    getTopNRooms = async () => {
        let self = this;
        let dateTime = moment().utc().format("YYYY-MM-DD HH:mm:ss");
        let rooms = { Query: "", Duration: 30, TimeZone: self.state.selectedTimeZone, Time: dateTime, IsScheduleRequired: true };
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

            this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} is unauthorized!`, severityLevel: SeverityLevel.Warning });
            this.setState({ authorized: false });
            return [];
        }
        else if (res.status === 200) {
            let response = await res.json();
            this.setState({ topFiveRooms: response });
        }
        else {
            this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
            this.appInsights.trackTrace({ message: `'TopNRoomsAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
        }
    }

    /**
     * Filter rooms as per user input.
     * @param inputValue Input string.
     */
    filterRooms = async (inputValue: string) => {
        let self = this;
        let dateTime = moment().utc().format("YYYY-MM-DD HH:mm:ss");

        if (inputValue) {
            let rooms = { Query: inputValue, Duration: self.state.selectedDuration.value, TimeZone: self.state.selectedTimeZone, Time: dateTime, IsScheduleRequired: true };
            const res = await fetch("/api/MeetingApi/SearchRoomAsync", {
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

                this.setState({ authorized: false });
                this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} is unauthorized!`, severityLevel: SeverityLevel.Warning });
                return [];
            }
            else if (res.status === 200) {
                let response = await res.json();
                return response;
            }
            else {
                this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                this.appInsights.trackTrace({ message: `'SearchRoomAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
            }
        }
    }

    /** Render duration dropdown. */
    renderDurationDropdown() {
        return (
            <div style={{ width: "inherit" }}>
                <Select options={this.state.duration}
                    theme={this.themeColor}
                    value={this.state.selectedDuration}
                    onChange={this.handleDurationChange}
                    components={{ IndicatorSeparator: () => null }}
                    styles={this.themeStyle}
                />
            </div>
        )
    }

    /** Render rooms dropdown. */
    renderRoomsDropdown() {
        const Option = (props: any) => {
            return (
                <components.Option {...props}>
                    <Text content={props.data.label} style={{ width: "auto", fontWeight: "bold" }} /><br />
                    <Text content={props.data.sublabel + " | "} /><Text content={props.data.Status} style={{ color: props.data.Status === "Available" ? "#92C353" : "#E74C3C" }} />
                </components.Option>
            );
        };

        return (
            <div style={{ width: "inherit" }}>
                <AsyncSelect
                    isDisabled={this.state.selectedDuration !== null && this.state.selectedTimeZone !== null ? false : true}
                    defaultOptions={this.state.topFiveRooms}
                    styles={this.themeStyle}
                    placeholder={this.state.resourceStrings.SearchRoomDropdownPlaceholder}
                    components={{ Option, IndicatorSeparator: () => null }}
                    loadOptions={this.promiseOptions.bind(this)}
                    value={this.state.selectedRoom}
                    onChange={this.handleRoomChange}
                    theme={this.themeColor}
                />
            </div>
        )
    }

    /**
     * Handles asynchronous operation for searching room.
     * @param inputValue Input string.
     */
    promiseOptions = (inputValue: string) =>
        new Promise(async resolve => {
            resolve(this.filterRooms(inputValue));
        });

    /**
     * Event called after selecting room.
     * @param optionSelected Selected room.
     */
    handleRoomChange = (optionSelected: any) => {
        this.setState({ selectedRoom: optionSelected, showMessage: false })
    }

    /**
     * Event called after selecting duration.
     * @param optionSelected Selected duration.
     */
    handleDurationChange = (optionSelected: any) => {
        this.setState({ selectedDuration: optionSelected });
    }

    /**
     * Event called after user selects timezone.
     * @param event Dropdown sythetic event object.
     * @param data Data props for dropdown.
     */
    handleTimezonSelectionChange = (event: React.SyntheticEvent<HTMLElement>, data?: any) => {
        this.setState({ selectedTimeZone: data.value, showMessage: false });
        let tzResult = this.state.supportedTimeZones.find(function (tz) { return tz === data.value });

        if (tzResult) {
            this.setState({ selectedTimeZone: data.value, showMessage: false });
            this.saveUserTimeZone(data.value);
            this.getTopNRooms();
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

    /** Submit selection and create meeting. */
    submit = async () => {
        this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} initiated meeting creation` });
        if (this.state.selectedDuration !== null && this.state.selectedRoom !== null) {
            let selectedRoom = this.state.selectedRoom;
            if (selectedRoom.Status === Constants.Available) {
                this.setState({ loading: true });
                let dateTime = moment().utc().format("YYYY-MM-DD HH:mm:ss");
                let meeting = { RoomName: selectedRoom.RoomName, RoomEmail: selectedRoom.RowKey, BuildingName: selectedRoom.BuildingName, Duration: parseInt(this.state.selectedDuration.value), TimeZone: this.state.selectedTimeZone, Time: dateTime, IsFavourite: false };
                const res = await fetch("/api/MeetingApi/CreateMeetingAsync", {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                        "Authorization": "Bearer " + this.token
                    },
                    body: JSON.stringify(meeting)
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

                    this.setState({ authorized: false, loading: false });
                }
                else if (res.status === 200) {
                    this.appInsights.trackEvent({ name: `Meeting created` }, { User: this.userObjectIdentifier, Room: selectedRoom.RowKey });
                    let response = await res.json();
                    if (response !== null) {
                        let toBot = { MeetingId: response.id, WebLink: response.webLink, RoomName: selectedRoom.RoomName, RoomEmail: selectedRoom.RowKey, BuildingName: selectedRoom.BuildingName, StartDateTime: response.start.timeZone, EndDateTime: response.end.timeZone, Text: "meeting from task module", isFavourite: false, replyTo: this.replyTo, BuildingEmail: selectedRoom.PartitionKey };
                        microsoftTeams.tasks.submitTask(toBot);
                    }
                    else {
                        this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                    }
                }
                else {
                    this.appInsights.trackTrace({ message: `'CreateMeetingAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
                    this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                }
            }
            else {
                this.setMessage(this.state.resourceStrings.RoomUnavailable, Constants.ErrorMessageRedColor, false);
            }
        }
        else {
            this.setMessage(this.state.resourceStrings.SelectDurationRoom, Constants.ErrorMessageRedColor, false);
        }
    }

    /** Show validation error message. */
    showError() {
        if (this.state.showMessage === true) {
            return (
                <Text error content={this.state.message} />
            );
        }
        else {
            return (<Text error content="" />);
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
            if (self.state.authorized === true) {
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
                                <Text weight="bold" content={self.state.resourceStrings.MeetingLength} />
                            </Flex>
                            <Flex gap="gap.small">
                                <Flex.Item grow>
                                    {self.renderDurationDropdown()}
                                </Flex.Item>
                            </Flex>
                            <Flex style={{ marginTop: "1rem" }}>
                                <Text weight="bold" content={self.state.resourceStrings.SearchRoom} />
                            </Flex>
                            <Flex gap="gap.small">
                                <Flex.Item grow>
                                    {self.renderRoomsDropdown()}
                                </Flex.Item>
                            </Flex>
                            <div className="footer">
                                <Flex gap="gap.small">
                                    <Flex.Item grow>
                                        {self.showError()}
                                    </Flex.Item>
                                    <Button loading={self.state.loading} primary disabled={self.state.selectedRoom === null || (self.state.selectedRoom === null && self.state.selectedDuration === null) || self.state.loading === true} onClick={() => self.submit()} content={self.state.resourceStrings.BookRoom} />
                                </Flex>
                            </div>
                        </div>
                    </Provider>
                );
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

export default withAITracking(reactPlugin, OtherRoom);