import { BotCommand, IBotParameters } from "../helpers/botCommand";
import { Utils } from "../helpers/utils";
const rawInfoCard = require("../adaptiveCards/info.json");

export class InfoCommand extends BotCommand {
  public name: string = InfoCommand.name;

  constructor() {
    super();
    this.matchPatterns = [/^\s*info\s*/];
  }

  validateParameters(parameters: any): boolean {
    if (!parameters.userData) {
      throw new Error(`Command "info" failed: missing input "userData"`);
    }
    return true;
  }

  async run(parameters: IBotParameters): Promise<void> {
    this.validateParameters(parameters);
    const card = Utils.renderAdaptiveCard(rawInfoCard, parameters.userData);
    await parameters.context.sendActivity({ attachments: [card] });
  }
}
