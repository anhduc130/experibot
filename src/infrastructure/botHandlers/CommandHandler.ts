import { ActionTypes, CardAction, CardFactory, MessageFactory, TeamsInfo, TurnContext } from "botbuilder";
import { IDependencies } from "../BotActivityHandler";
import { helpCard } from "../cards/helpCard";
import { identityCard } from "../cards/identityCard";
import { openTaskModuleCard } from "../cards/openTaskModuleCard";
import { refreshCard } from "../cards/refreshCard";
import { ActivityHandler } from "./ActivityHandler";
import { BubbleDemoHandler } from "./BubbleDemoHandler";
import { PaymentInMeetingHandler } from "./PaymentInMeetingHandler";
import { TargetedBubbleHandler } from "./TargetedBubbleHandler";

interface ISuggestedActionList {
  [key: string]: {
    message: string;
    actions: CardAction[];
  };
}

const Actions: { [key: string]: string } = {
  // SIGNIN: "signin",
  // SHOW_TASK_MODULE: "show task module",
  // SHOW_BUBBLE: "show bubble",
  // SHOW_TARGETED_BUBBLE: "show targeted bubble",
  // SHOW_BUBBLE_CLOSE: "show closing bubble",
  // SHOW_REFRESH: "show refresh",
  SHOW_SUGGESTED_ACTIONS: "IT Support",
  A1: "Finance",
  A2: "Human Resources",
  A3: "R&D",
  // START_ACTIVITY: "start activity",
  // CONFIRM_ANONYMOUS_IDENTITY: "confirm identity",
  // MEETING_IS_DONE: "meeting is done",
  HELP: "help",
  // MONITOR: "monitor participants",
};
const COMPLETE_ACTIVITY = "complete activity";
const COMPLETE_PAYMENT = "complete payment";

const suggestedActionList: ISuggestedActionList = {
  [Actions.SHOW_SUGGESTED_ACTIONS.toLowerCase()]: {
    message: "Hi! Thanks for contacting IT Support. How can I help you today?",
    actions: [
      {
        type: "imBack",
        title: "Equipment request",
        value: "Equipment request",
      },
      {
        type: "imBack",
        title: "Software licenses",
        value: "Software licenses",
      },
      {
        type: "imBack",
        title: "Access rights",
        value: "Access rights",
      },
    ],
  },
  "access rights": {
    message: "Thanks. What would you like to do?",
    actions: [
      {
        type: "imBack",
        title: "I need more access",
        value: "I need more access",
      },
      {
        type: "imBack",
        title: "I request access for someone else",
        value: "I request access for someone else",
      },
    ],
  },
  "i need more access": {
    message: "Thanks. To request more access, please visit: //myaccess. You will need an approval from your manager.",
    actions: [],
  },
  yes: {
    message: "Great! Thank you for using our bot service. 😄",
    actions: [],
  },
  doesItHelp: {
    message: "Does it help?",
    actions: [
      {
        type: "imBack",
        title: "Yes",
        value: "Yes",
      },
      {
        type: "imBack",
        title: "No, please connect me with an agent",
        value: "No, please connect me with an agent",
      },
    ],
  },
  doesItHelpEdit: {
    message: "[Edit] Does it help?",
    actions: [
      {
        type: "imBack",
        title: "[Edit] Yes",
        value: "[Edit] Yes",
      },
      {
        type: "imBack",
        title: "[Edit] No, please connect me with an agent",
        value: "[Edit] No, please connect me with an agent",
      },
    ],
  },
};

export class CommandHandler {
  static Actions = Actions;

  constructor(
    private deps: IDependencies,
    private activityHandler: ActivityHandler,
    private bubbleDemoHandler: BubbleDemoHandler,
    private targetedBubbleDemoHandler: TargetedBubbleHandler,
    private paymentHandler: PaymentInMeetingHandler
  ) {}

  async handleCommand(command: string, context: TurnContext) {
    switch (command) {
      case Actions.HELP:
        await this.helpActivityAsync(context, command);
        break;
      case Actions.SHOW_REFRESH:
        await this.showRefreshCardAsync(context);
        break;
      case Actions.SHOW_TASK_MODULE:
        await this.showTaskModuleAsync(context);
        break;
      case Actions.SHOW_TARGETED_BUBBLE:
        await this.targetedBubbleDemoHandler.showTargetedBubbleAsync(context);
        break;
      case Actions.SHOW_BUBBLE:
        await this.bubbleDemoHandler.showBubbleAsync(context);
        break;
      case Actions.SHOW_BUBBLE_CLOSE:
        await this.bubbleDemoHandler.showClosingBubbleAsync(context);
        break;
      case Actions.SHOW_SUGGESTED_ACTIONS.toLowerCase():
      case "access rights":
      case "i need more access":
      case "yes":
        await this.showSuggestedActionsAsync(context, command);
        break;
      case Actions.CONFIRM_ANONYMOUS_IDENTITY:
        await this.confirmAnonymousIdentityAsync(context);
        break;
      case Actions.START_ACTIVITY:
        await this.activityHandler.startActivityAsync(context);
        break;
      case COMPLETE_ACTIVITY:
        await this.activityHandler.completeActivityAsync(context);
        break;
      case Actions.MEETING_IS_DONE:
        await this.meetingIsDoneAsync(context);
        break;
      case Actions.MONITOR:
        await this.monitorAsync(context);
      default:
      // await this.helpActivityAsync(context, command);
    }
  }

  async meetingIsDoneAsync(context: TurnContext) {
    const replyActivity = MessageFactory.text("Meeting is done!"); // this could be an adaptive card instead
    const img = encodeURIComponent("https://i.imgur.com/RbCKrf8.gif");
    const url = `${process.env.BaseUrl}/bubble/meeting-is-done.html?message=${img}`;
    const encodedUrl = encodeURIComponent(url as string);
    const height = 500;
    const width = 400;
    replyActivity.channelData = {
      notification: {
        alertInMeeting: true,
        externalResourceUrl: `https://teams.microsoft.com/l/bubble/${process.env.BotId}?url=${encodedUrl}&height=${height}&width=${width}&title=Meeting%20is%20finished&completionBotId=${process.env.BotId}`,
      },
    };
    await context.sendActivity(replyActivity);
  }

  private async monitorAsync(context: TurnContext) {
    const participants = await TeamsInfo.getMembers(context);
    const names = participants.map((p) => p.name).join(", ");
    console.log(names);
    setTimeout(() => this.monitorAsync(context).then(() => {}), 10000);
  }

  private async confirmAnonymousIdentityAsync(context: TurnContext) {
    const userId = context.activity.from.id;
    const msa = this.deps.identityManager.getIdentityFromUserId(userId) || "No MSA mapping found";
    const card = CardFactory.adaptiveCard(identityCard(msa, userId));
    await context.sendActivity({ attachments: [card] });
  }

  private async showTaskModuleAsync(context: TurnContext) {
    const card = CardFactory.adaptiveCard(openTaskModuleCard());
    this.deps.logger.log(`From: ${JSON.stringify(context.activity.from, null, 2)}`);
    await context.sendActivity({
      attachments: [card],
      suggestedActions: {
        actions: [
          {
            title: "green",
            type: "imBack",
            value: "green",
          },
        ],
        to: [context.activity.from.id],
      },
    });
  }

  private async showRefreshCardAsync(context: TurnContext) {
    const member = await TeamsInfo.getMember(context, context.activity.from.id);
    const members = await TeamsInfo.getMembers(context);
    const ids = members.map((member) => member.id);
    const card = CardFactory.adaptiveCard(refreshCard("Initial message", member.name, ids));
    await context.sendActivity({ attachments: [card] });
  }

  private async showSuggestedActionsAsync(context: TurnContext, command: string) {
    const buildActivity = (key: string) => {
      const { message, actions } = suggestedActionList[key];
      const to = [context.activity.from.id];
      const activity = MessageFactory.text(message);
      activity.suggestedActions = { actions, to };
      return activity;
    };

    const activities = [buildActivity(command)];

    await context.sendActivities(activities);

    if (command === "i need more access") {
      const response = await context.sendActivity(buildActivity("doesItHelp"));

      await new Promise((resolve) => setTimeout(resolve, 5000));

      const newActivity = buildActivity("doesItHelpEdit");
      newActivity.id = response?.id;
      await context.updateActivity(newActivity);
    }
  }

  private async helpActivityAsync(context: TurnContext, text: string) {
    const card = CardFactory.adaptiveCard(helpCard(Actions, text));
    await context.sendActivity({
      attachments: [card],
    });
  }
}
