import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import Model = formattingSettings.Model;
import SimpleCard = formattingSettings.SimpleCard;
declare class LayoutCardSettings extends SimpleCard {
    name: string;
    displayName: string;
    slices: (formattingSettings.NumUpDown | formattingSettings.ItemDropdown)[];
}
declare class AppearanceCardSettings extends SimpleCard {
    name: string;
    displayName: string;
    slices: (formattingSettings.ToggleSwitch | formattingSettings.ColorPicker)[];
}
declare class LinesCardSettings extends SimpleCard {
    name: string;
    displayName: string;
    slices: (formattingSettings.ToggleSwitch | formattingSettings.NumUpDown | formattingSettings.ItemDropdown | formattingSettings.ColorPicker)[];
}
declare class NodesCardSettings extends SimpleCard {
    name: string;
    displayName: string;
    slices: (formattingSettings.ToggleSwitch | formattingSettings.NumUpDown | formattingSettings.ItemDropdown | formattingSettings.ColorPicker | formattingSettings.TextInput)[];
}
declare class LevelsCardSettings extends SimpleCard {
    name: string;
    displayName: string;
    slices: (formattingSettings.ToggleSwitch | formattingSettings.ColorPicker)[];
}
declare class ControlsCardSettings extends SimpleCard {
    name: string;
    displayName: string;
    slices: (formattingSettings.ToggleSwitch | formattingSettings.NumUpDown | formattingSettings.ItemDropdown)[];
}
declare class TableCardSettings extends SimpleCard {
    name: string;
    displayName: string;
    slices: (formattingSettings.ToggleSwitch | formattingSettings.NumUpDown)[];
}
export declare class VisualFormattingSettingsModel extends Model {
    layout: LayoutCardSettings;
    appearance: AppearanceCardSettings;
    lines: LinesCardSettings;
    nodes: NodesCardSettings;
    levels: LevelsCardSettings;
    controls: ControlsCardSettings;
    table: TableCardSettings;
    cards: (LayoutCardSettings | AppearanceCardSettings | LinesCardSettings | NodesCardSettings | LevelsCardSettings | ControlsCardSettings | TableCardSettings)[];
}
export {};
