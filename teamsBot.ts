import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
  MessageFactory,
} from "botbuilder";
import rawWelcomeCard from "./adaptiveCards/welcome.json";
import rawLearnCard from "./adaptiveCards/learn.json";
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";

export interface DataInterface {
  likeCount: number;
}

const cereInstanceURL = "http://127.0.0.1:3000";

export class TeamsBot extends TeamsActivityHandler {
  // record the likeCount
  likeCountObj: { likeCount: number };

  constructor() {
    super();

    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      try {
        const url = new URL(cereInstanceURL + "/api/teams/chat");
        const raw = JSON.stringify({
          question: context.activity.text,
          history: [],
        });

        const response = await fetch(url, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: raw,
        });
        const aiResponse = await response.json();
        const replyText = aiResponse.text;
        await context.sendActivity(MessageFactory.text(replyText, replyText));
        // By calling next() you ensure that the next BotHandler is run.
      } catch (error) {
        console.log(error);
      }
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;

      const url = new URL(cereInstanceURL + "/api/settings");

      try {
        const response = await fetch(url, {
          method: "GET",
          headers: {
            "Content-Type": "application/json",
          },
        });
        const settings = await response.json();
        const welcomeText = settings.data.opening;
        for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
          if (membersAdded[cnt].id !== context.activity.recipient.id) {
            await context.sendActivity(
              MessageFactory.text(welcomeText, welcomeText)
            );
          }
        }
      } catch (error) {
        console.log(error);
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
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "userlike") {
      this.likeCountObj.likeCount++;
      const card = AdaptiveCards.declare<DataInterface>(rawLearnCard).render(
        this.likeCountObj
      );
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200, type: undefined, value: undefined };
    }
  }
}
