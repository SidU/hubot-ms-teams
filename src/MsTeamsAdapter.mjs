import { Adapter } from 'hubot'
import EventEmitter from 'node:events'
import {
    MessageFactory,
    CardFactory,
    TextFormatTypes,
    TurnContext
} from 'botbuilder'
import { MessageActivity } from '@microsoft/teams.api'

const CONTENT_LENGTH_LIMIT = 2_000

class MsTeamsAdapter extends Adapter {
    #client
    #activityHandler
    constructor(robot, activityHandler = new EventEmitter(), client = new EventEmitter()) {
        super(robot)
        this.#activityHandler = activityHandler
        this.#client = client
        this.#client.onTurnError = this.#onTurnError
        this.conversationReferences = {}
    }
    async #onTurnError(context, error) {
        this.robot.logger.info(`[onTurnError] ${error} ${JSON.stringify(context)}`)
        await context.sendTraceActivity('onTurnError trace', `${error}`, 'https://www.botframework.com/schemas/error', 'TurnError')
        await context.sendActivity('The bot encountered an error.')
    }

    async send(envelope, ...strings) {
        // Handle messageRoom calls where envelope only has room property
        if (envelope.room && !envelope.user) {
            return await this.sendToRoom(envelope.room, ...strings)
        }
        // Handle regular send calls with user.message
        const responses = await this.sendWithDelegate(envelope.user.message, envelope, ...strings)
        this.emit('send', envelope, responses)
        return responses
    }

    async sendToRoom(room, ...strings) {
        const serviceUrl = process.env.TEAMS_BOT_SERVICE_URL || `https://smba.trafficmanager.net/amer/${process.env.TEAMS_BOT_TENANT_ID}/`

        // conversationReferences is keyed by the conversation id.
        const conversationReference = this.conversationReferences[room.channelData.channel.id]
        if (!conversationReference) {
            this.robot.logger.error(`No conversation reference found for room, creating a new one: ${room.channelData.channel.id}`)
            const conversationParameters = {
                isGroup: true,
                bot: { id: process.env.TEAMS_BOT_APP_ID, name: this.robot.name},
                serviceUrl: serviceUrl,
                channelData: room.channelData,
                activity: MessageFactory.text(strings.join('\n')),
                tenantId: process.env.TEAMS_BOT_TENANT_ID
            }
            // The conversationReferences key is the conversation id.
            // botAppId, channelId, serviceUrl, audience, conversationParameters, logic
            await this.#client.createConversationAsync(process.env.TEAMS_BOT_APP_ID,
                'msteams', // channel here means which platform is this in. Slack, MSTeams, etc.
                serviceUrl,
                null, // audience
                conversationParameters,
                async turnContext => {
                    this.conversationReferences[turnContext.activity.conversation.id] = turnContext.activity.conversation
                    await turnContext.sendActivity(strings.join('\n'))
                }
            )
            this.robot.logger.debug(`Created new conversation reference for room: ${JSON.stringify(room, null, 2)}`)
            return []
        }
        
        const responses = []
        await this.#client.continueConversation(conversationReference, async (context) => {
            for await (let message of strings) {
                let teamsMessage = MessageFactory.text(message, message)
                let card = null

                teamsMessage.textFormat = TextFormatTypes.Markdown
                if (/<\/(.*)>/.test(message)) {
                    teamsMessage.textFormat = TextFormatTypes.Xml
                }
                
                try {
                    card = JSON.parse(message)
                    teamsMessage = {
                        attachments: [ CardFactory.adaptiveCard(card) ]
                    }
                } catch(e) {
                    this.robot.logger.debug(`message isn't a card: ${e}`)
                }
                
                try {
                    const response = await context.sendActivity(teamsMessage)
                    if (response) {
                        responses.push(response)
                    }
                } catch (e) {
                    if(e.statusCode && e.statusCode === 401){
                        this.robot.logger.error(`${this.robot.name}: Unauthorized, check TEAMS_BOT_APP_ID, TEAMS_BOT_CLIENT_SECRET, TEAMS_BOT_APP_TYPE, and TEAMS_BOT_TENANT_ID`)
                    } else {
                        this.robot.logger.error(`${this.robot.name}: ${e}`)
                    }
                }
            }
        })
        
        this.emit('send', { room }, responses)
        return responses
    }
    async reply(envelope, ...strings) {
        const responses = await this.sendWithDelegate(envelope.user.message, envelope, ...strings)
        this.emit('reply', envelope, responses)
        return responses
    }
    async sendWithDelegate(delegate, envelope, ...strings) {
        const responses = []
        for await (let message of strings) {
            let teamsMessage = MessageFactory.text(message, message)
            let card = null

            teamsMessage.textFormat = TextFormatTypes.Markdown
            if (/<\/(.*)>/.test(message)) {
                teamsMessage.textFormat = TextFormatTypes.Xml
            }
            
            try {
                card = JSON.parse(message)
                teamsMessage = {
                    attachments: [ CardFactory.adaptiveCard(card) ]
                }
            } catch(e) {
                this.robot.logger.debug(`message isn't a card: ${e}`)
            }
            try {
                const response = await delegate.sendActivity(teamsMessage)
                if (response) {
                    responses.push(response)
                }
            } catch (e) {
                if(e.statusCode && e.statusCode === 401){
                    this.robot.logger.error(`${this.robot.name}: Unauthorized, check TEAMS_BOT_APP_ID, TEAMS_BOT_CLIENT_SECRET, TEAMS_BOT_APP_TYPE, and TEAMS_BOT_TENANT_ID`)
                } else {
                    this.robot.logger.error(`${this.robot.name}: ${e}`)
                }
            }
        }
        return responses
    }
    #escapeRegex(value) {
        return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    }

    #ensureMentionPrefix(text, robotName) {
        if (!robotName) {
            return text?.trim() ?? text
        }
        const trimmed = (text ?? '').trim()
        if (!trimmed) {
            return `@${robotName}`
        }
        const mentionPattern = new RegExp(`^@${this.#escapeRegex(robotName)}\\b`, 'i')
        if (mentionPattern.test(trimmed)) {
            return trimmed.replace(mentionPattern, `@${robotName}`)
        }
        const barePattern = new RegExp(`^${this.#escapeRegex(robotName)}\\b`, 'i')
        if (barePattern.test(trimmed)) {
            return trimmed.replace(barePattern, `@${robotName}`)
        }
        return `@${robotName} ${trimmed}`
    }

    #normalizeIncomingActivity(activity) {
        const robotName = (this.robot.alias == false ? undefined : this.robot.alias) ?? this.robot.name
        if (!activity || typeof activity !== 'object') {
            return activity
        }

        const messageActivity = MessageActivity.from(activity)
        const botAccountId = messageActivity.recipient?.id

        if (messageActivity.text) {
            if (botAccountId) {
                messageActivity.stripMentionsText({ accountId: botAccountId, tagOnly: true })
            } else {
                messageActivity.stripMentionsText({ tagOnly: true })
            }
        }

        let normalizedText = (messageActivity.text ?? '')
            .replace(/^\r?\n/, '')
            .replace(/\\n$/, '')
            .trim()

        let mentionDetected = false

        if (normalizedText.includes('<at>')) {
            mentionDetected = true
            normalizedText = normalizedText.replace(/<at>(.*?)<\/at>/gi, '$1')
        }

        const isPersonal = activity?.conversation?.conversationType === 'personal'
        const botMention = botAccountId ? messageActivity.getAccountMention(botAccountId) : null
        mentionDetected = mentionDetected || !!botMention

        if ((isPersonal || mentionDetected) && robotName && normalizedText.length > 0) {
            normalizedText = this.#ensureMentionPrefix(normalizedText, robotName)
        }

        activity.text = normalizedText
        return activity
    }
    async run() {
        this.robot.router.use(async (req, res, next) => {
            this.robot.logger.debug(`url: ${req.url}`)
            this.robot.logger.debug(`headers: ${JSON.stringify(req.headers)}`)
            this.robot.logger.debug(`body: ${JSON.stringify(req.body)}`)
            next()
        })
        this.robot.router.post(['/', '/api/messages'], async (req, res)=>{
            req.body = this.#normalizeIncomingActivity(req.body)

            try {
                await this.#client.process(req, res, async context => {
                    // Store conversation reference for messageRoom functionality
                    const conversationReference = TurnContext.getConversationReference(context.activity)
                    this.conversationReferences[context.activity.conversation.id] = conversationReference
                    await this.#activityHandler.run(context)
                    res.status(200).send('ok')
                })
            } catch (e) {
                this.robot.logger.error(e)
                res.status(500).send('service error')
            }
        })
        this.robot.server.on('upgrade', async (req, socket, head) => {
            this.robot.logger.info('upgrading to websockets')
            await this.#client.process(req, socket, head, (context) => this.#activityHandler.run(context));
        })
        this.emit('connected', this)
        this.robot.logger.info(`${MsTeamsAdapter.name} adapter is running as @${this.robot.name}.`)
    }
    close () {
        this.robot.logger.info(`${MsTeamsAdapter.name} adapter is closing.`)
        this.emit('disconnected')
    }
}
export default MsTeamsAdapter
export {
    MsTeamsAdapter
}  