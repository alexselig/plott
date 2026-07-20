#!/usr/bin/env node
/**
 * Generate the Office Add-in manifest for a given host origin. All variants share
 * one stable add-in Id (identity), differing only in the URLs they point at.
 *
 *   node scripts/make-manifest.mjs [hostOrigin] [outFile]
 *
 * Examples:
 *   node scripts/make-manifest.mjs                              # localhost dev -> public/manifest.xml
 *   node scripts/make-manifest.mjs https://alexselig.github.io/plott public/manifest.ghpages.xml
 *
 * The host origin must be HTTPS (Office requires it) and must serve this app's
 * static export, so that `<host>/addin/` renders the task pane.
 */
import { writeFileSync } from "node:fs";

// Stable identity — do NOT regenerate; changing it makes Office treat it as a new add-in.
const ADDIN_ID = "2876137d-d349-4ed1-9493-5b31797709ca";

const host = (process.argv[2] || "https://localhost:3010").replace(/\/+$/, "");
const out = process.argv[3] || "public/manifest.xml";

const manifest = `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>${ADDIN_ID}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Alex Selig</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Plott" />
  <Description DefaultValue="Design presentation-ready charts and insert them onto your slides — then reopen and restyle a chart that's already on a slide." />
  <IconUrl DefaultValue="${host}/addin/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="${host}/addin/icon-80.png" />
  <SupportUrl DefaultValue="${host}/" />
  <AppDomains>
    <AppDomain>${host}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Presentation" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${host}/addin/" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <FunctionFile resid="Taskpane.Url" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Plott.Group">
                <Label resid="Plott.GroupLabel" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16" />
                  <bt:Image size="32" resid="Icon.32" />
                  <bt:Image size="80" resid="Icon.80" />
                </Icon>
                <Control xsi:type="Button" id="Plott.OpenPane">
                  <Label resid="Plott.OpenLabel" />
                  <Supertip>
                    <Title resid="Plott.OpenLabel" />
                    <Description resid="Plott.OpenDesc" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16" />
                    <bt:Image size="32" resid="Icon.32" />
                    <bt:Image size="80" resid="Icon.80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>PlottTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${host}/addin/icon-16.png" />
        <bt:Image id="Icon.32" DefaultValue="${host}/addin/icon-32.png" />
        <bt:Image id="Icon.80" DefaultValue="${host}/addin/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${host}/addin/" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Plott.GroupLabel" DefaultValue="Plott" />
        <bt:String id="Plott.OpenLabel" DefaultValue="Chart builder" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Plott.OpenDesc" DefaultValue="Design a chart and insert it, or restyle a chart already on the slide." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;

writeFileSync(out, manifest);
console.log(`Wrote ${out} for host ${host}`);
