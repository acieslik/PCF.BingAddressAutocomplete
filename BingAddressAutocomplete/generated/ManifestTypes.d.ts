/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    value: ComponentFramework.PropertyTypes.StringProperty;
    bingapikey: ComponentFramework.PropertyTypes.StringProperty;
    street: ComponentFramework.PropertyTypes.StringProperty;
    city: ComponentFramework.PropertyTypes.StringProperty;
    county: ComponentFramework.PropertyTypes.StringProperty;
    state: ComponentFramework.PropertyTypes.StringProperty;
    zipcode: ComponentFramework.PropertyTypes.StringProperty;
    country: ComponentFramework.PropertyTypes.StringProperty;
}
export interface IOutputs {
    value?: string;
    street?: string;
    city?: string;
    county?: string;
    state?: string;
    zipcode?: string;
    country?: string;
}
