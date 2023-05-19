import { TurnContext } from "botbuilder";
import { BotCommand, IBotParameters } from "../helpers/botCommand";

export class SSOCommand extends BotCommand {
  public name: string = SSOCommand.name;

  public operationWithSSOToken!: ((
    context: TurnContext,
    ssoToken: string,
    userState: any
  ) => Promise<void> | undefined);
  
  validateParameters(parameters: IBotParameters): boolean {
    if (!parameters.ssoDialog) {
      throw new Error(`SSOCommand failed: missing input "ssoDialog".`);
    }
    if (!parameters.context) {
      throw new Error(`SSOCommand failed: missing input "context".`);
    }
    if (!parameters.dialogState) {
      throw new Error(`SSOCommand failed: missing input "dialogState".`);
    }
    return true;
  }

  async run(parameters: IBotParameters): Promise<void> {
    this.validateParameters(parameters);
    const ssoDialog = parameters.ssoDialog;
    ssoDialog.setSSOOperation(this.operationWithSSOToken);
    await ssoDialog.run(parameters.context, parameters.dialogState);
  }
}
  