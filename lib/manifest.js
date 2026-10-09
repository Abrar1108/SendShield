// Generates the Outlook add-in manifest. Served live from /manifest.xml so the ribbon text
// edited in the dashboard is included. Outlook reads the manifest only when it is (re)uploaded,
// so ribbon changes bump the version's last number and need a re-upload to take effect.
export const ADDIN_ID = "cb06eb0b-cc2d-4d26-a426-600bd698fa54";
export const APP_VERSION = "6.1.0";

function xml(text) {
    return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

export function manifestVersion(revision = 0) {
    const [a = 0, b = 0, c = 0] = APP_VERSION.split(".").map(Number);
    return `${a}.${b}.${c}.${Math.max(0, revision | 0)}`;
}

export function renderManifest({ baseUrl, messages, revision = 0, displayName = "SendShield", addinId = ADDIN_ID }) {
    const u = xml(baseUrl);
    return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
          xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
          xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
          xsi:type="MailApp">
  <Id>${xml(addinId)}</Id>
  <Version>${manifestVersion(revision)}</Version>
  <ProviderName>Toluna</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="${xml(displayName)}" />
  <Description DefaultValue="Asks users to confirm external recipients before an email leaves Toluna." />
  <IconUrl DefaultValue="${u}/addin/assets/sendshield-64.png" />
  <HighResolutionIconUrl DefaultValue="${u}/addin/assets/sendshield-128.png" />
  <SupportUrl DefaultValue="${u}/" />
  <AppDomains>
    <AppDomain>${u}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.3" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="${u}/addin/taskpane.html"/>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Requirements>
        <bt:Sets DefaultMinVersion="1.12">
          <bt:Set Name="Mailbox" />
        </bt:Sets>
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <Runtimes>
            <Runtime resid="WebViewRuntime.Url">
              <Override type="javascript" resid="JSRuntime.Url"/>
            </Runtime>
          </Runtimes>
          <DesktopFormFactor>
            <FunctionFile resid="Commands.Url" />
            <ExtensionPoint xsi:type="MessageComposeCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="msgComposeGroup">
                  <Label resid="GroupLabel" />
                  <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                    <Label resid="TaskpaneButton.Label" />
                    <Supertip>
                      <Title resid="TaskpaneButton.Label" />
                      <Description resid="TaskpaneButton.Tooltip" />
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16" />
                      <bt:Image size="32" resid="Icon.32x32" />
                      <bt:Image size="80" resid="Icon.80x80" />
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url" />
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
            <ExtensionPoint xsi:type="LaunchEvent">
              <LaunchEvents>
                <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock" />
                <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onRecipientsChangedHandler" />
              </LaunchEvents>
              <SourceLocation resid="WebViewRuntime.Url"/>
            </ExtensionPoint>
          </DesktopFormFactor>
        </Host>
      </Hosts>
      <Resources>
        <bt:Images>
          <bt:Image id="Icon.16x16" DefaultValue="${u}/addin/assets/sendshield-16.png"/>
          <bt:Image id="Icon.32x32" DefaultValue="${u}/addin/assets/sendshield-32.png"/>
          <bt:Image id="Icon.80x80" DefaultValue="${u}/addin/assets/sendshield-80.png"/>
        </bt:Images>
        <bt:Urls>
          <bt:Url id="Commands.Url" DefaultValue="${u}/addin/handler.html" />
          <bt:Url id="Taskpane.Url" DefaultValue="${u}/addin/taskpane.html" />
          <bt:Url id="WebViewRuntime.Url" DefaultValue="${u}/addin/handler.html" />
          <bt:Url id="JSRuntime.Url" DefaultValue="${u}/addin/handler.js" />
        </bt:Urls>
        <bt:ShortStrings>
          <bt:String id="GroupLabel" DefaultValue="${xml(messages.ribbonGroup)}"/>
          <bt:String id="TaskpaneButton.Label" DefaultValue="${xml(messages.ribbonButton)}"/>
        </bt:ShortStrings>
        <bt:LongStrings>
          <bt:String id="TaskpaneButton.Tooltip" DefaultValue="${xml(messages.ribbonTooltip)}"/>
        </bt:LongStrings>
      </Resources>
    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>
`;
}
