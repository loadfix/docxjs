import { OpenXmlElement } from "./dom";

export interface WmlInstructionText extends OpenXmlElement {
    text: string;
}

// Legacy (pre-SDT) form-field data carried on a <w:fldChar fldCharType="begin">
// via a <w:ffData> child. Only the subset needed to render a read-only
// <input>/<select> is retained. All string fields are attacker-controlled
// DOCX data and must only reach the DOM via setAttribute / textContent.
export interface FormFieldData {
    // "text" (FORMTEXT), "checkbox" (FORMCHECKBOX), "dropdown" (FORMDROPDOWN)
    formFieldType?: 'text' | 'checkbox' | 'dropdown';
    // FORMTEXT: default textbox value (<w:textInput><w:default w:val="…"/>).
    defaultText?: string;
    // FORMTEXT: max length in chars (<w:textInput><w:maxLength w:val="…"/>).
    maxLength?: number;
    // FORMCHECKBOX: checked state (<w:checkBox><w:checked/> or
    // <w:default w:val="1"/>).
    checked?: boolean;
    // FORMDROPDOWN: list entries (<w:ddList><w:listEntry w:val="…"/>).
    ddItems?: string[];
    // FORMDROPDOWN: index of the selected entry (<w:ddList><w:default w:val="…"/>).
    ddDefault?: number;
}

export interface WmlFieldChar extends OpenXmlElement {
    charType: 'begin' | 'end' | 'separate' | string;
    lock: boolean;
    dirty?: boolean;
    // Present only on fldCharType="begin" runs that carry a <w:ffData>
    // child (legacy form fields).
    ffData?: FormFieldData;
}

export interface WmlFieldSimple extends OpenXmlElement {
    instruction: string;
    lock: boolean;
    dirty: boolean;
}