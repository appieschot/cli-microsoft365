import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';

import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { SpoOperation } from './SpoOperation';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  skipReycleBin?: boolean;
  fromRecycleBin?: boolean;
  wait?: boolean;
  confirm?: boolean;
}

class SpoSiteClassicRemoveCommand extends SpoCommand {
  private formDigest?: string;
  private formDigestExpiresAt?: Date;
  private accessToken?: string;
  private dots?: string;
  private timeout?: NodeJS.Timer;

  public get name(): string {
    return commands.LIST_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.url = (!(!args.options.url)).toString();
    telemetryProps.skipReycleBin = (!(!args.options.skipReycleBin)).toString();
    telemetryProps.fromRecycleBin = (!(!args.options.fromRecycleBin)).toString();
    telemetryProps.wait = (!(!args.options.wait)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {

    const removeSite: () => void = (): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((accessToken: string): Promise<void> => {
          if (this.debug) {
            cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest for tenant admin at ${auth.site.url}...`);
          }

          this.accessToken = accessToken;

          return this.ensureFormDigest(cmd);
        })
        .then((): request.RequestPromise => {
          if (this.verbose) {
            cmd.log(`Deleting site collection ${args.options.url}...`);
          }

          var body = `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}">
            <Actions>
              <ObjectPath Id="55" ObjectPathId="54"/>
              <ObjectPath Id="57" ObjectPathId="56"/>
              <Query Id="58" ObjectPathId="54">
                <Query SelectAllProperties="true">
                  <Properties/>
                </Query>
              </Query>
              <Query Id="59" ObjectPathId="56">
                <Query SelectAllProperties="false">
                  <Properties>
                    <Property Name="IsComplete" ScalarProperty="true"/>
                    <Property Name="PollingInterval" ScalarProperty="true"/>
                  </Properties>
                </Query>
              </Query>
            </Actions>
            <ObjectPaths>
              <Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/>
              <Method Id="56" ParentId="54" Name="RemoveSite">
                <Parameters>
                  <Parameter Type="String">${Utils.escapeXml(args.options.url)}</Parameter>
                </Parameters>
              </Method>
            </ObjectPaths>
          </Request>`; 

          // todo implement the fromRecycleBin Auto when --skiprecyclebin is implemented
          if(args.options.fromRecycleBin){
            if (this.verbose) {
              cmd.log(`Permanitly remove the site from the Recycle Bin `);
            }

            body = `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}">
            <Actions>
              <ObjectPath Id="63" ObjectPathId="62"/>
              <ObjectPath Id="65" ObjectPathId="64"/>
              <Query Id="66" ObjectPathId="64">
                <Query SelectAllProperties="false">
                  <Properties>
                    <Property Name="IsComplete" ScalarProperty="true"/>
                    <Property Name="PollingInterval" ScalarProperty="true"/>
                  </Properties>
                </Query>
              </Query>
            </Actions>
            <ObjectPaths>
              <Constructor Id="62" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/>
              <Method Id="64" ParentId="62" Name="RemoveDeletedSite">
                <Parameters>
                  <Parameter Type="String">${Utils.escapeXml(args.options.url)}</Parameter>
                </Parameters>
              </Method>
            </ObjectPaths>
          </Request>`; 
          }

          const requestOptions: any = {
            url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              'X-RequestDigest': this.formDigest
            }),
            body: body
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.post(requestOptions);
        })
        .then((res: string): Promise<void> => {
          return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
            if (this.debug) {
              cmd.log('Response:');
              cmd.log(res);
              cmd.log('');
            }

            const json: ClientSvcResponse = JSON.parse(res);
            const response: ClientSvcResponseContents = json[0];
            if (response.ErrorInfo) {
              reject(response.ErrorInfo.ErrorMessage);
            }
            else {
              const operation: SpoOperation = json[json.length - 1];
              let isComplete: boolean = operation.IsComplete;

              if (!args.options.wait || isComplete) {
                resolve();
                return;
              }

              this.timeout = setTimeout(() => {
                this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), resolve, reject, this.accessToken as string, cmd);
              }, operation.PollingInterval);
            }
          });
        })
        .then((): void => {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
    }

    if (args.options.confirm) {
      removeSite();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the site ${args.options.url}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeSite();
        }
      });
    }
  }

  private ensureFormDigest(cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const now: Date = new Date();
      if (this.formDigest &&
        now < (this.formDigestExpiresAt as Date)) {
        if (this.debug) {
          cmd.log('Existing form digest still valid');
        }

        resolve();
        return;
      }

      this
        .getRequestDigest(cmd, this.debug)
        .then((res: ContextInfo): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          this.formDigest = res.FormDigestValue;
          this.formDigestExpiresAt = new Date();
          this.formDigestExpiresAt.setSeconds(this.formDigestExpiresAt.getSeconds() + res.FormDigestTimeoutSeconds - 5);

          resolve();
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private waitUntilFinished(operationId: string, resolve: () => void, reject: (error: any) => void, accessToken: string, cmd: CommandInstance): void {
    this
      .ensureFormDigest(cmd)
      .then((): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Checking if operation ${operationId} completed...`);
        }

        if (!this.debug && this.verbose) {
          this.dots += '.';
          process.stdout.write(`\r${this.dots}`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': this.formDigest
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="${operationId.replace(/\\n/g, '&#xA;').replace(/"/g, '')}" /></ObjectPaths></Request>`
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          reject(response.ErrorInfo.ErrorMessage);
        }
        else {
          const operation: SpoOperation = json[json.length - 1];
          let isComplete: boolean = operation.IsComplete;
          if (isComplete) {
            if (this.verbose) {
              process.stdout.write('\n');
            }

            resolve();
            return;
          }

          this.timeout = setTimeout(() => {
            this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), resolve, reject, accessToken, cmd);
          }, operation.PollingInterval);
        }
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'url of the site to remove'
      },
      {
        option: '--skipRecycleBin',
        description: 'set to directly remove the site without moving it to the Recycle Bin'
      },
      {
        option: '--fromRecycleBin',
        description: 'set to remove the site from the Recycle Bin'
      },
      {
        option: '--wait',
        description: 'Wait for the site to be removed before completing the command'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the list'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.url) {
        return 'Required parameter url missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
    using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:
  
    To remove a classic site, you have to first connect to a tenant admin
    site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
        
    Deleting and creating classic site collections is by default asynchronous
    and depending on the current state of Office 365, might take up to few
    minutes. If you're building a script with steps that require the site to be
    fully provisioned, you should use the ${chalk.blue('--wait')} flag. When
    using this flag, the ${chalk.blue(this.getCommandName())} command will keep
    running until it received confirmation from Office 365 that the site
    has been fully provisioned.

  Examples:
  
    Remove the site based on URL, and place it in the recycle bin
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.SITE_CLASSIC_REMOVE} --url https://contoso.sharepoint.com/sites/project-x 

    Remove the site based on URL permanitly 
      ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
        ${chalk.grey(config.delimiter)} ${commands.SITE_CLASSIC_REMOVE} --url https://contoso.sharepoint.com/sites/project-x --skipRecycleBin

    Remove the site based on URL from the recycle bin
      ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
        ${chalk.grey(config.delimiter)} ${commands.SITE_CLASSIC_REMOVE} --url https://contoso.sharepoint.com/sites/project-x --fromRecycleBin

    Remove the site based on URL permanitly and wait for completion 
      ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
        ${chalk.grey(config.delimiter)} ${commands.SITE_CLASSIC_REMOVE} --url https://contoso.sharepoint.com/sites/project-x --wait --skipRecycleBin
      `);
  }
}

module.exports = new SpoSiteClassicRemoveCommand();