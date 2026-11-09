# Teams web API notes

The Teams web app uses internal APIs. Microsoft does not document them, and they can change.
These notes are facts that we checked on a real account. Update them when something changes.

## Sign-in and tokens

* The page holds one account per tenant (home account and guest accounts). Use only the primary account (`localStorage` key `…PrimaryUserId`).
* The skype token comes from the answer of `POST teams.microsoft.com/api/authsvc/v1.0/authz`. The tenant is the `tid` of the bearer token of that request.
* The service addresses are in `localStorage` key `tmp.auth.v1.<userId>.Discover.DISCOVER-REGION-GTM` (`chatServiceAfd`, `chatSvcAggAfd`, `middleTier`, …).
* The chat service needs `Authentication: skypetoken=…`. The profile lookup needs only `Authorization: Bearer` (audience `https://api.spaces.skype.com`). With both headers it answers 400.

## Chats

* List: `GET {chatServiceAfd}/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=500`.
* Kind: `threadProperties.productThreadType` (`OneToOneChat`, `Chat`, `Meeting`, `TeamsStandardChannel`, `TeamsTeam`, …). IDs that start with `48:` are system streams.
* Deleted teams stay in the list with `threadProperties.isdeleted`. Channels that you left have a `memberExpirationTime` in the past and answer 404 "soft-deleted". Meeting chats with an expired membership stay readable.
* Unread: `lastMessage.id` is greater than the first field of `properties.consumptionhorizon`.
* A 1:1 chat has no title. Its ID is `19:<guid>_<guid>@unq.gbl.spaces`. After a tenant move, old chats hold your old user ID, so also compare display names.
* A channel's `threadProperties.spaceThreadTopic` holds the team name and `topicThreadTopic` the channel name, so the team name needs no extra request. The CSA service (team/favorite grouping) needs a `chatsvcagg.teams.microsoft.com` token that Teams does not mint for the home tenant while the tool drives it, so the tool does not use CSA.
* Users from other tenants are not found by `fetchShortProfile`. Use the names on their messages (`imdisplayname`).

## Messages

* Read: `GET …/conversations/{id}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=N&startTime=1`. Older pages: `_metadata.backwardLink`. Follow it only on the chat-service host.
* Send: `POST …/conversations/{id}/messages` with `content`, `messagetype: RichText/Html`, `contenttype: text`, `imdisplayname` (required) and `clientmessageid`. The answer field `OriginalArrivalTime` is the message ID.
* Reply in a channel thread: `…/conversations/{channelId};messageid={rootId}/messages`.
* Scheduled send: `POST …/users/ME/drafts` with `draftType: ScheduledDraft` and `draftDetails.sendAt`.
* `properties.emotions` is a JSON string or a list. The key `follow` is a thread subscription, not a reaction.
* Mentions: `<span itemtype="http://schema.skype.com/Mention" itemscope itemid="N">` in the content, plus `properties.mentions` (JSON string).

Sources: our own tests, [msteams-mcp](https://github.com/m0nkmaster/msteams-mcp) and [teams-api](https://github.com/Maxim-Mazurok/teams-api).
