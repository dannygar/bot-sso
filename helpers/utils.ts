import { CardFactory, Attachment } from "botbuilder";
import ACData = require("adaptivecards-templating");

export class Utils {
  // Bind AdaptiveCard with data
  static renderAdaptiveCard(rawCardTemplate: any, dataObj?: any): Attachment {
    const cardTemplate = new ACData.Template(rawCardTemplate);
    const cardWithData = cardTemplate.expand({ 
      $root: {
        "userData": dataObj
      } 
    });
    const card = CardFactory.adaptiveCard(cardWithData);
    return card;
  }

  static toPascalCase(text: string): string {
    return `${text}`
      .toLowerCase()
      .replace(new RegExp(/[-_]+/, 'g'), ' ')
      .replace(new RegExp(/[^\w\s]/, 'g'), '')
      .replace(
        new RegExp(/\s+(.)(\w*)/, 'g'),
        ($1, $2, $3) => `${$2.toUpperCase() + $3}`
      )
      .replace(new RegExp(/\w/), s => s.toUpperCase());
  }  
}
