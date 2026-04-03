"use strict";

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class MarkerSettingsCard extends FormattingSettingsCard {
    tagSize = new formattingSettings.NumUpDown({
        name: "tagSize",
        displayName: "Tag Size",
        value: 1.5,
        options: {
            minValue: {
                type: powerbi.visuals.ValidatorType.Min,
                value: 0.1
            },
            maxValue: {
                type: powerbi.visuals.ValidatorType.Max,
                value: 6
            }
        }
    });

    name: string = "markerSettings";
    displayName: string = "Marker Settings";
    slices: Array<FormattingSettingsSlice> = [this.tagSize];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    markerSettingsCard = new MarkerSettingsCard();
    cards = [this.markerSettingsCard];
}