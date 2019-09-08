import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory } from "botbuilder";
import { MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder-teams";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import { ITaskModuleResult, IMessagingExtensionActionRequest } from "botbuilder-teams-messagingextensions";
// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/timelyMessageExtension/config.html")
@PreventIframe("/timelyMessageExtension/action.html")
export default class TimelyMessageExtension implements IMessagingExtensionMiddlewareProcessor {



    public async onFetchTask(context: TurnContext, value: IMessagingExtensionActionRequest): Promise<MessagingExtensionResult | ITaskModuleResult> {



        return Promise.resolve<ITaskModuleResult>({
            type: "continue",
            value: {
                title: "Input form",
                url: `https://${process.env.HOSTNAME}/timelyMessageExtension/action.html`
            }
        });


    }


    // handle action response in here
    // See documentation for `MessagingExtensionResult` for details
    public async onSubmitAction(context: TurnContext, value: IMessagingExtensionActionRequest): Promise<MessagingExtensionResult> {


        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: value.data.email
                    },
                    {
                        type: "Image",
                        url: `https://randomuser.me/api/portraits/thumb/women/${Math.round(Math.random() * 100)}.jpg`
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.0"
            });
        return Promise.resolve({
            type: "result",
            attachmentLayout: "list",
            attachments: [card]
        } as MessagingExtensionResult);
    }



}
