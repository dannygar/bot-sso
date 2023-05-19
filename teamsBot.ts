import {
  TeamsActivityHandler,
  TurnContext,
  SigninStateVerificationQuery,
  BotState,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
  MemoryStorage,
  ConversationState,
  UserState,
  StatePropertyAccessor,
} from "botbuilder";
import { Utils } from "./helpers/utils";
import { SSODialog } from "./helpers/ssoDialog";
import { SkillsHelper } from "./helpers/skillsHelper";
import { commands } from "./commands";
import { SSOCommand } from "./commands/ssoCommand";
import { SSOSignIn } from "./commands/ssoSignIn";
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const rawInfoCard = require("./adaptiveCards/info.json");

const CONVERSATION_DATA_PROPERTY = 'conversationData';
const USER_PROFILE_PROPERTY = 'userProfile';

export class TeamsBot extends TeamsActivityHandler {
  private conversationState: BotState;
  private userState: BotState;
  private conversationDataAccessor: StatePropertyAccessor<any>;
  private userProfileAccessor: StatePropertyAccessor<any>;
  private dialog: SSODialog;
  private dialogState: any;
  private dataObject: any = {};

  constructor(conversationState: ConversationState, userState: UserState) {
    super();

    if (!conversationState) {
      throw new Error('[TeamsBot]: Missing parameter. conversationState is required');
    }
    if (!userState) {
        throw new Error('[TeamsBot]: Missing parameter. userState is required');
    }

    this.conversationState = conversationState;
    this.userState = userState;

    // Create conversation and user state with in-memory storage provider.
    this.dialog = new SSODialog(new MemoryStorage(), this.userState);
    this.dialogState = this.conversationState.createProperty("DialogState");

    // Create the state property accessors for the conversation data and user profile.
    this.conversationDataAccessor = conversationState.createProperty(CONVERSATION_DATA_PROPERTY);
    this.userProfileAccessor = userState.createProperty(USER_PROFILE_PROPERTY);


    // Greet the user once the chat is created
    this.onConversationUpdate(async (context: TurnContext) => {
      if (context.activity.membersAdded) {
          for (const member of context.activity.membersAdded) {
              if (member.id !== context.activity.recipient.id) {
                // Trigger SSO sign-in
                commands.find((c) => c.name === SSOCommand.name)?.run(
                  {
                    context: context,
                    ssoDialog: this.dialog,
                    dialogState: this.dialogState,
                  }
                );
              }
          }
      }
    });        

    // Reply to user messages
    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      let txt = context.activity.text;
      // remove the mention of this bot
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Get the state properties from the turn context.
      const userProfile = await this.userProfileAccessor.get(context, {});
      // If the 'DidBotWelcomedUser' does not exist (first time ever for a user), set the default to false.
      const conversationData = await this.conversationDataAccessor.get(context, {didUserSignedIn: false});

      // Check if the user wants to sign in
      if (!conversationData.didUserSignedIn) {
        // Trigger SSO sign-in
        const ssoCommand = commands.find((c) => c.name === SSOSignIn.name);
        if (ssoCommand) {
          ssoCommand.run(
            {
              context: context,
              ssoDialog: this.dialog,
              dialogState: this.dialogState,
              userData: userProfile,
            }
          );
          await this.conversationDataAccessor.set(context, {didUserSignedIn: true});
        }
      } 

      // Trigger command by IM text
      const isCommandExecuted = await SkillsHelper.triggerSkillDialog(txt, {
        context: context,
        ssoDialog: this.dialog,
        dialogState: this.dialogState,
      });

      // If no command is executed, show the welcome adaptive card
      if (!isCommandExecuted) {
        const card = Utils.renderAdaptiveCard(rawWelcomeCard);
        await context.sendActivity({ attachments: [card] });
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // Sends welcome messages to conversation members when they join the conversation.
    // Messages are only sent to conversation members who aren't the bot.
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      if (membersAdded === undefined) {
        return await next();
      }
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = Utils.renderAdaptiveCard(rawWelcomeCard);
          await context.sendActivity({ attachments: [card] });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    // The verb "get info about" is sent from the Adaptive Card defined in adaptiveCards/info.json
    if (invokeValue.action.verb === "info") {
      const card = Utils.renderAdaptiveCard(rawInfoCard, this.dataObject);
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [card],
      });
    }
    return { statusCode: 200, type: '', value: {}, };
  }

  async run(context: TurnContext) {
    await super.run(context);

    // Save any state changes. The load happened during the execution of the Dialog.
    await this.conversationState.saveChanges(context, false);
    await this.userState.saveChanges(context, false);
  }

  async handleTeamsSigninVerifyState(
    context: TurnContext,
    query: SigninStateVerificationQuery
  ) {
    console.log(
      "Running dialog with signin/verifystate from an Invoke Activity."
    );
    await this.dialog.run(context, this.dialogState);
  }

  async handleTeamsSigninTokenExchange(
    context: TurnContext,
    query: SigninStateVerificationQuery
  ) {
    await this.dialog.run(context, this.dialogState);
  }

  async onSignInInvoke(context: TurnContext) {
    await this.dialog.run(context, this.dialogState);
  }
}
