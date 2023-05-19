import { CardFactory, TurnContext } from "botbuilder";
import { SSOCommand } from "./ssoCommand";
import ApiGraph from "../api/apiGraph";

export class ShowUserProfile extends SSOCommand {
  public name: string = ShowUserProfile.name;

  constructor() {
    super();
    this.matchPatterns = [/^\s*show\s*/];
    this.operationWithSSOToken = this.showUserInfo;
  }


  async showUserInfo(context: TurnContext, ssoToken: string, userState: any) {
    await context.sendActivity("Retrieving user information from Microsoft Graph ...");

    // Create a Graph client
    const graphApi = new ApiGraph(ssoToken);

    // Get user profile
    const me = await graphApi.getPersonAsync();
    if (me) {
      await context.sendActivity(
        `You're logged in as ${me.displayName} (${me.userPrincipalName})${
          me.jobTitle ? `; your job title is: ${me.jobTitle}` : ""
        }.`
      );

      // Get user picture
      const userPicture = await graphApi.getUserPhotoAsync();

      // show user picture
      const card = CardFactory.thumbnailCard(
        "User Picture",
        CardFactory.images([userPicture])
      );
      await context.sendActivity({ attachments: [card] });
    } else {
      await context.sendActivity(
        "Could not retrieve profile information from Microsoft Graph."
      );
    }
  }
}
