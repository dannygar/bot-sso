import { TurnContext } from "botbuilder";
import { SSOCommand } from "./ssoCommand";

export class SSOSignIn extends SSOCommand {
  public name: string = SSOSignIn.name;

  constructor() {
    super();
    this.matchPatterns = [/^\s*sign\s*/];
    this.operationWithSSOToken = this.signIn;
  }

  async signIn(context: TurnContext, ssoToken: string, userState: any) {
    userState.ssoToken = ssoToken;
    await context.sendActivity(`Your SSO token is: ${ssoToken}`);
  }
}
