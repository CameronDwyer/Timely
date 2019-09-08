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
                url: `https://${process.env.HOSTNAME}/timelyMessageExtension/action.html`,
                width: "medium",
                height: "large"
            }
        });


    }


    // handle action response in here
    // See documentation for `MessagingExtensionResult` for details
    public async onSubmitAction(context: TurnContext, value: IMessagingExtensionActionRequest): Promise<MessagingExtensionResult> {
        // Create array of facts with the property names needed for the Adaptive Card
        const facts: Array<{title: string, value: string}> = value.data.timezonesConversions.map(item => {
            return {
                title: item.locationName,
                value: item.time
            }
        });

        const card = CardFactory.adaptiveCard(
            {
                "type": "AdaptiveCard",
                "version": "1.0",
                "body": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "Image",
                                        "altText": "",
                                        "url": `https://${process.env.HOSTNAME}/assets/drop-pin-logo.png`,
                                        "size": "Small"
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "FactSet",
                                        "facts": facts,
                                    }
                                ]
                            }
                        ]
                    }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json"
            });
        return Promise.resolve({
            type: "result",
            attachmentLayout: "list",
            attachments: [card]
        } as MessagingExtensionResult);
    }



}
