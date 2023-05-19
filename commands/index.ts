import { BotCommand } from "../helpers/botCommand";
import { InfoCommand } from "./info";
import { ShowUserProfile } from "./showUserProfile";
import { SSOSignIn } from "./ssoSignIn";
import { WelcomeCommand } from "./welcome";

export const commands: BotCommand[] = [
  new WelcomeCommand(),
  new SSOSignIn(),
  new ShowUserProfile(),
  new InfoCommand(),
];
