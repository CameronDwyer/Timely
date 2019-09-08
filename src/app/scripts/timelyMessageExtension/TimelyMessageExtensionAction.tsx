import * as React from "react";
import {
    PrimaryButton,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    Input,
    TeamsThemeContext,
    getContext
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import * as moment from "moment-timezone";
import Select from "react-select";

/**
 * State for the TimelyMessageExtensionAction React component
 */
export interface ITimelyMessageExtensionActionState extends ITeamsBaseComponentState {
    baseLocation: string;
    baseTime: string;
    timezonesConversions: Array<{locationName: string, time: string}>;
}

/**
 * Properties for the TimelyMessageExtensionAction React component
 */
export interface ITimelyMessageExtensionActionProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the Timely Message Extension Task Module page
 */
export class TimelyMessageExtensionAction extends TeamsBaseComponent<ITimelyMessageExtensionActionProps, ITimelyMessageExtensionActionState> {
    private selectedZones: Array<{locationName: string}> = new Array();

    public componentWillMount() {
        this.initaliseLocations();
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize(),
            timezonesConversions: new Array(),
            baseTime: this.currentDateTimeInLocalDateTimeFormat()
        });
        microsoftTeams.initialize();
        microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes, weights } = font;
        const styles = {
            header: { ...sizes.title2, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall },
            selector: { minHeight: 180 },
            list: {...sizes.base, marginLeft: 0, paddingLeft: 0}
        };

        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>Convert from</div>
                        </PanelHeader>
                        <PanelBody>
                            <div style={styles.section}>
                                <Select
                                    value={{value: this.state.baseLocation, label: this.state.baseLocation}}
                                    options={this.getLocationOptions()}
                                    onChange={(loc) => {
                                        if (loc.label !== this.state.baseLocation) {
                                            this.setState( {
                                                baseLocation: loc.value
                                            });
                                        }
                                        this.redrawSelectedTimezoneTable();
                                    }}
                                    isMulti={false}
                                    maxMenuHeight={120}
                                />
                                <Input
                                    type="datetime-local"
                                    autoFocus
                                    label="Date / Time"
                                    errorLabel={!this.state.baseTime ? "This value is required" : undefined}
                                    value={this.state.baseTime}
                                    onChange={(e) => {
                                        this.setState({
                                            baseTime: e.target.value
                                        });
                                        this.redrawSelectedTimezoneTable();
                                    }}
                                    required />
                            </div>
                        </PanelBody>
                        <PanelHeader>
                            <div style={styles.header}>Convert to</div>
                        </PanelHeader>
                        <PanelBody>
                            <div style={{...styles.section, ...styles.selector}} >
                                <Select
                                    options={this.getLocationOptions()}
                                    onChange={this.newTargetLocationSelected.bind(this)}
                                    isMulti={false}
                                    maxMenuHeight={120}
                                />
                                <ul style={styles.list}>
                                    {this.state.timezonesConversions.map((item, index) => (
                                        <li>{item.locationName}   {item.time}</li>
                                    ))}
                                </ul>
                            </div>
                        </PanelBody>
                        <PanelFooter>
                            <div style={styles.section}>
                                <PrimaryButton onClick={() => {
                                    microsoftTeams.tasks.submitTask({
                                        timezonesConversions: this.state.timezonesConversions
                                    });
                                }}>OK</PrimaryButton>
                            </div>
                            <div style={styles.footer}>
                                (C) Copyright Camtoso
                            </div>
                        </PanelFooter>
                    </Panel>
                </Surface>
             </TeamsThemeContext.Provider>
        );
    }

    private getLocationOptions() {
        const options: Array<{value: string, label: string}> = new Array();
        const availableZones = moment.tz.names();
        availableZones.forEach(option =>
                options.push({value: option, label: option})
            );
        return options;
    }

    private initaliseLocations() {
        const defaultZone = moment.tz.guess();
        this.setState({baseLocation: defaultZone});
    }

    private newTargetLocationSelected(newLocation) {
        this.selectedZones.push({locationName: newLocation.value});
        this.redrawSelectedTimezoneTable();
    }

    private currentDateTimeInLocalDateTimeFormat() {
        const now: Date = new Date();
        const utcString: string = now.toISOString().substring(0, 19);
        const year: number = now.getFullYear();
        const month: number = now.getMonth() + 1;
        const day: number = now.getDate();
        const hour: number = now.getHours();
        const minute: number = now.getMinutes();
        const localDatetime: string = year + "-" +
            (month < 10 ? "0" + month.toString() : month) + "-" +
            (day < 10 ? "0" + day.toString() : day) + "T" +
            (hour < 10 ? "0" + hour.toString() : hour) + ":" +
            (minute < 10 ? "0" + minute.toString() : minute) +
            utcString.substring(16, 19);
        return localDatetime;
    }

     private redrawSelectedTimezoneTable() {
        // rebuild timezone conversion array
        const timezones: Array<{locationName: string, time: string}> = new Array();
        timezones.push({locationName: this.state.baseLocation, time: moment.tz(this.state.baseTime, this.state.baseLocation).tz(this.state.baseLocation).format("llll")});
        this.selectedZones.forEach( entry => {
            timezones.push({locationName: entry.locationName, time: moment.tz(this.state.baseTime, this.state.baseLocation).tz(entry.locationName).format("llll")});
        });
        this.setState({timezonesConversions: timezones});
    }
}
