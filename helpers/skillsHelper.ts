import { commands } from "../commands";
import { IBotParameters } from "./botCommand";
import { Utils } from "./utils";

export class SkillsHelper {
  static async triggerSkillDialog(userInput: string, parameters: IBotParameters): Promise<boolean> {
    for (let command of commands) {
      const matchText = command.expressionMatchesText(userInput);
      if (matchText) {
        if (typeof matchText !== "boolean") // RegExpExecArray 
        {
          const parameter: string = userInput.replace(matchText[0], "").trim();
          parameters.userData = Utils.toPascalCase(parameter);
        }
        await command.run(parameters);
        return true;
      }
    }
    return false;
  }
}
