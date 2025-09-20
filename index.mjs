import { MsTeamsAdapter } from './src/MsTeamsAdapter.mjs'
import { ActivityHandler } from 'botbuilder'
import { TextMessage, User } from 'hubot'
import { BotBuilderPlugin } from '@microsoft/teams.botbuilder'
import { createRequire } from 'node:module'

const require = createRequire(import.meta.url)
const { Client } = require('@microsoft/teams.common/http')
const { ConsoleLogger } = require('@microsoft/teams.common/logging')

const defaultMessageMapper = context => {

    // Try to simplify the message structure to reduce the memory footprint.
    // I think the problem is that TextMessage (via Message) gets it's room from user.room.
    // And user.room is the entire context.activity object which is huuuuuge.
    // This is an opportunity to feedback into the Hubot Message structure design.
    // Seems to be the interplay between TextMessage (Message) and the envelope in send().
    const activity = context.activity
    const sharedActivity = {
        text: activity.text,
        textFormat: activity.textFormat,
        attachments: activity.attachments,
        type: activity.type,
        timestamp: activity.timestamp,
        localTimestamp: activity.localTimestamp,
        id: activity.id,
        channelId: activity.channelId,
        serviceUrl: activity.serviceUrl,
        from: activity.from,
        conversation: activity.conversation,
        recipient: activity.recipient,
        entities: activity.entities,
        channelData: activity.channelData,
        locale: activity.locale,
        localTimezone: activity.localTimezone,
        rawTimestamp: activity.rawTimestamp,
        rawLocalTimestamp: activity.rawLocalTimestamp,
        callerId: activity.callerId
    }

    const message = new TextMessage(new User(context.activity.from.id, {
        name: context.activity.from.name,
        room: new Proxy(sharedActivity, {
            get(target, prop) {
                return target[prop]
            },
            set(target, prop, value) {
                target[prop] = value
                return true
            }
        }),
        message: context  // this is what the code uses to send messages to MS Bot Service Platform
    }), context.activity.text, context.activity.id)
    return message
}

class HubotActivityHandler extends ActivityHandler {
    #robot = null
    #messageMapper = null
    constructor(robot, messageMapper = defaultMessageMapper) {
        super()
        this.#messageMapper = messageMapper ?? defaultMessageMapper
        this.#robot = robot
        this.onMessage(async (context, next) => {
            await this.#robot.receive(this.#messageMapper(context))
            await next()
        })
    }
}
export {
    HubotActivityHandler
}
export default {
    async use(robot) {
        robot.config = {
            TEAMS_BOT_CLIENT_SECRET: process.env.TEAMS_BOT_CLIENT_SECRET ?? null,
            TEAMS_BOT_TENANT_ID: process.env.TEAMS_BOT_TENANT_ID ?? null,
            TEAMS_BOT_APP_ID: process.env.TEAMS_BOT_APP_ID ?? null,
            TEAMS_BOT_APP_TYPE: process.env.TEAMS_BOT_APP_TYPE ?? null
        }
        const activityHandler = new HubotActivityHandler(robot)
        const plugin = new BotBuilderPlugin({ handler: activityHandler })
        const tenantId = (process.env.TEAMS_BOT_APP_TYPE ?? '').toLowerCase() === 'singletenant'
            ? process.env.TEAMS_BOT_TENANT_ID
            : undefined

        plugin.logger = new ConsoleLogger('@hubot-friends/hubot-ms-teams')
        plugin.client = new Client({
            headers: {
                'User-Agent': '@hubot-friends/hubot-ms-teams'
            }
        })
        plugin.manifest = {
            name: {
                short: robot.name,
                full: robot.name
            }
        }
        plugin.credentials = {
            clientId: process.env.TEAMS_BOT_APP_ID,
            clientSecret: process.env.TEAMS_BOT_CLIENT_SECRET,
            tenantId
        }
        plugin.onInit()

        const adapter = new MsTeamsAdapter(robot, activityHandler, plugin.adapter)
        return adapter
    }
}
