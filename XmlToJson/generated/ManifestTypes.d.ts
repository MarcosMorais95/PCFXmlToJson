/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    Reset: ComponentFramework.PropertyTypes.WholeNumberProperty;
    AllowMultiple: ComponentFramework.PropertyTypes.TwoOptionsProperty;
    IsSchemaVisible: ComponentFramework.PropertyTypes.TwoOptionsProperty;
}
export interface IOutputs {
    jsonResult?: string;
}
