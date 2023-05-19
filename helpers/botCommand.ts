import { TurnContext } from "botbuilder";
import { SSODialog } from "./ssoDialog";

export type PredicateFunc<T> = (v: T) => boolean;
export type MatchTerm = string | RegExp | PredicateFunc<string>;

export interface IBotParameters {
  context: TurnContext;
  ssoDialog: SSODialog;
  dialogState: any;
  userData?: any;
}

export abstract class BotCommand {
  public matchPatterns: MatchTerm[] = [];
  abstract name: string;

  abstract run(parameters: IBotParameters): Promise<void>;

  public validateParameters(parameters: IBotParameters): boolean {
    return true;
  }

  public expressionMatchesText(userInput: string): RegExpExecArray | boolean {
    let matchResult: RegExpExecArray | boolean | null;
    for (const pattern of this.matchPatterns) {
      if (typeof pattern === "string") {
        matchResult = new RegExp(pattern).exec(userInput);
      } else if (pattern instanceof RegExp) {
        matchResult = pattern.exec(userInput);
      } else {
        matchResult = pattern(userInput);
      }
      if (matchResult) {
        return matchResult;
      }
    }
    return false;
  }
}

