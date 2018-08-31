import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import GraphCommand from '../../GraphCommand';
import { CommandOption, CommandValidate } from '../../../../Command';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;

}

interface Options extends GlobalOptions {
  classifications: string;
  defaultClassification: string;
  usageGuidelinesUrl?: string;
}

class GraphO365SiteClassificationEnableCommand extends GraphCommand {
  public get name(): string {
    return `${commands.SITECLASSIFICATION_ENABLE}`;
  }

  public get description(): string {
    return 'Enables site classification configuration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.classifications = args.options.classifications;
    telemetryProps.defaultClassification = args.options.defaultClassification;
    telemetryProps.usageGuidelinesUrl = args.options.usageGuidelinesUrl;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        // ToDo: Handle nullable usageguidelinesurl prop & check templateId
        //
        const requestOptions: any = {
          url: `${auth.service.resource}/beta/settings`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          }),
          json: true,
          body: JSON.parse(
            `{
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [{
                "name": "ClassificationList",
                "value": "${args.options.classifications}"
              }, {
                "name": "DefaultClassification",
                "value": "${args.options.defaultClassification}"
              }, {
                "name": "UsageGuidelinesUrl",
                "value": "${args.options.usageGuidelinesUrl ? null : args.options.usageGuidelinesUrl || ''}"
              }]
            }`
          )
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        // if (res.value.length == 0) {
        //   cb(new CommandError('Site classification is not enabled.'));
        //   return;
        // }

        // const unifiedGroupSetting: DirectorySetting[] = res.value.filter((directorySetting: DirectorySetting): boolean => {
        //   return directorySetting.displayName === 'Group.Unified';
        // });

        // if (unifiedGroupSetting == null || unifiedGroupSetting.length == 0) {
        //   cb(new CommandError("Missing DirectorySettingTemplate for \"Group.Unified\""));
        //   return;
        // }

        // const siteClassificationsSettings: SiteClassificationSettings = new SiteClassificationSettings();

        // // Get the classification list
        // const classificationList: DirectorySettingValue[] = unifiedGroupSetting[0].values.filter((directorySetting: DirectorySettingValue): boolean => {
        //   return directorySetting.name === 'ClassificationList';
        // });

        // siteClassificationsSettings.Classifications = [];
        // if (classificationList != null && classificationList.length > 0) {
        //   siteClassificationsSettings.Classifications = classificationList[0].value.split(',');
        // }

        // // Get the UsageGuidancelinesUrl
        // const guidanceUrl: DirectorySettingValue[] = unifiedGroupSetting[0].values.filter((directorySetting: DirectorySettingValue): boolean => {
        //   return directorySetting.name === 'UsageGuidelinesUrl';
        // });

        // siteClassificationsSettings.UsageGuidelinesUrl = "";
        // if (guidanceUrl != null && guidanceUrl.length > 0) {
        //   siteClassificationsSettings.UsageGuidelinesUrl = guidanceUrl[0].value;
        // }

        // // Get the DefaultClassification
        // const defaultClassification: DirectorySettingValue[] = unifiedGroupSetting[0].values.filter((directorySetting: DirectorySettingValue): boolean => {
        //   return directorySetting.name === 'DefaultClassification';
        // });

        // siteClassificationsSettings.DefaultClassification = "";
        // if (defaultClassification != null && defaultClassification.length > 0) {
        //   siteClassificationsSettings.DefaultClassification = defaultClassification[0].value;
        // }

        // cmd.log(siteClassificationsSettings);

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-c, --classifications',
        description: 'comma-separated list of classifications to enable in the tenant'
      },
      {
        option: '-d, --defaultClassification',
        description: 'classification to use by default'
      },
      {
        option: '--usageGuidelinesUrl',
        description: 'URL with additional information that should be displayed when choosing the classification for the given site'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.classifications) {
        return 'Required option classifications missing';
      }

      if (!args.options.defaultClassification) {
        return 'Required option defaultclassification missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to the Microsoft Graph
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.

    To set the Office 365 Tenant site classification, you have
    to first connect to the Microsoft Graph using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT}`)}.

  Examples:
  
    // Todo: Fill in the samples

  More information:

    SharePoint "modern" sites classification
      https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification
    `);
  }
}

module.exports = new GraphO365SiteClassificationEnableCommand();