/*
    <copyright file="constants.tsx" company="Microsoft Corporation">
    Copyright (c) Microsoft Corporation. All rights reserved.
    </copyright>
*/

export const Available = "Available";
export const Unavailable = "Unavailable";
export const ErrorMessageRedColor = "#E74C3C";
export const DefaultTheme = "default";
export const DarkTheme = "dark";
export const DurationArray = [{ label: "30 min", value: 30, key: 30 }, { label: "60 min", value: 60, key: 60 }, { label: "90 min", value: 90, key: 90 }];

//Theme style and colors
export const DefaultStyles = {
    control: (styles: any, { data, isDisabled, isFocused, isSelected }: any) => {
        return {
            ...styles,
            backgroundColor: "#F3F2F1",
            borderRight: isFocused ? 0 : 0,
            borderLeft: isFocused ? 0 : 0,
            borderTop: isFocused ? 0 : 0,
            boxShadow: isFocused ? 0 : 0,
            borderBottom: isFocused ? "2px solid #464775" : 0
        };
    },
    option: (styles: any, { data, isDisabled, isFocused, isSelected }: any) => {
        return {
            ...styles,
            backgroundColor: isDisabled ? null : isSelected ? "#6264A7" : isFocused ? "#F3F2F1" : null,
            color: isDisabled ? "#ccc" : isSelected ? "white" : isFocused ? "#252423" : "black",
            cursor: isDisabled ? "not-allowed" : "default",

            ":active": {
                ...styles[":active"],
                backgroundColor: !isDisabled && (isSelected ? "#6264A7" : isFocused ? "#F3F2F1" : "#F8F9F9"),
            },
        };
    }
};
export const DarkStyles = {
    control: (styles: any, { data, isDisabled, isFocused, isSelected }: any) => {
        return {
            ...styles,
            backgroundColor: "#454545",
            borderRight: isFocused ? 0 : 0,
            borderLeft: isFocused ? 0 : 0,
            borderTop: isFocused ? 0 : 0,
            boxShadow: isFocused ? 0 : 0,
            borderBottom: isFocused ? "2px solid #464775" : 0
        };
    },
    option: (styles: any, { data, isDisabled, isFocused, isSelected }: any) => {
        return {
            ...styles,
            backgroundColor: isDisabled ? null : isSelected ? "#6264A7" : isFocused ? "#F3F2F1" : null,
            color: isDisabled ? "#ccc" : isSelected ? "white" : isFocused ? "#252423" : "white",
            cursor: isDisabled ? "not-allowed" : "default",

            ":active": {
                ...styles[":active"],
                backgroundColor: !isDisabled && (isSelected ? "#6264A7" : isFocused ? "#F3F2F1" : "#808080"),
            },
        };
    }
};
export const ContrastStyles = {
    control: (styles: any, { data, isDisabled, isFocused, isSelected }: any) => {
        return {
            ...styles,
            backgroundColor: "black",
            borderRight: isFocused ? "2px solid white" : "2px solid white",
            borderLeft: isFocused ? "2px solid white" : "2px solid white",
            borderTop: isFocused ? "2px solid white" : "2px solid white",
            boxShadow: isFocused ? "1px 1px 1px 1px white" : "1px 1px 1px 1px white",
            borderBottom: "2px solid white"
        };
    },
    option: (styles: any, { data, isDisabled, isFocused, isSelected }: any) => {
        return {
            ...styles,
            backgroundColor: isDisabled ? null : isSelected ? "#FFFF01" : isFocused ? "#FFFF01" : null,
            color: isDisabled ? "#ccc" : isSelected ? "black" : isFocused ? "black" : "white",
            cursor: isDisabled ? "not-allowed" : "default",
            borderRight: isFocused ? "2px solid white" : "2px solid white",
            borderLeft: isFocused ? "2px solid white" : "2px solid white",
            ":active": {
                ...styles[":active"],
                backgroundColor: !isDisabled && (isSelected ? "#FFFF01" : isFocused ? "#FFFF01" : "black"),
            }
        };
    }
};
export const DefaultThemeColors = (theme: any) => {
    return {
        ...theme,
        borderRadius: 1,
        colors: {
            ...theme.colors,
            neutral0: "white",
            primary: "#6264A7",
            neutral5: "white",
            neutral80: "black",
            neutral50: "black"
        },
    }
};
export const DarkThemeColors = (theme: any) => {
    return {
        ...theme,
        borderRadius: 1,
        colors: {
            ...theme.colors,
            primary25: "#2d2c2c",
            neutral0: "#3b3a3a",
            primary: "#6264A7",
            neutral5: "#3b3a3a",
            neutral80: "white",
            neutral50: "white"
        },
    }
};
export const ContrastThemeColors = (theme: any) => {
    return {
        ...theme,
        borderRadius: 0,
        colors: {
            ...theme.colors,
            neutral0: "black",
            primary: "#FFFF01",
            neutral5: "white",
            neutral80: "#FFFF01",
            neutral50: "white",
            neutral10: "white",
            neutral15: "white"
        },
    }
};