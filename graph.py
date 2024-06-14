from configparser import SectionProxy
from typing import Optional
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.users_request_builder import UsersRequestBuilder

from msgraph.generated.chats.chats_request_builder import ChatsRequestBuilder
from msgraph.generated.chats.get_all_messages.get_all_messages_request_builder import (
    GetAllMessagesRequestBuilder,
)

from msgraph.generated.certificate_based_auth_configuration.certificate_based_auth_configuration_request_builder import (
    RequestConfiguration,
)

from msgraph.generated.models.team_collection_response import TeamCollectionResponse
from msgraph.generated.models.channel_collection_response import (
    ChannelCollectionResponse,
)

from msgraph.generated.models.chat_message_collection_response import (
    ChatMessageCollectionResponse,
)

# from kiota_abstractions.base_request_configuration import RequestConfiguration


class Graph:
    settings: SectionProxy
    client_credential: ClientSecretCredential
    app_client: GraphServiceClient

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings["clientId"]
        tenant_id = self.settings["tenantId"]
        client_secret = self.settings["clientSecret"]

        self.client_credential = ClientSecretCredential(
            tenant_id, client_id, client_secret
        )
        self.app_client = GraphServiceClient(self.client_credential)  # type: ignore

    async def get_app_only_token(self):
        graph_scope = "https://graph.microsoft.com/.default"
        access_token = await self.client_credential.get_token(graph_scope)
        return access_token.token

    async def get_users(self):
        query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            # Only request specific properties
            select=["displayName", "id", "mail"],
            # Get at most 25 results
            top=25,
            # Sort by display name
            orderby=["displayName"],
        )
        request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        users = await self.app_client.users.get(request_configuration=request_config)
        return users

    async def get_all_teams(self) -> Optional[TeamCollectionResponse]:
        result = await self.app_client.teams.get()

        return result

    async def get_all_channels_in_team(
        self, teamID: str
    ) -> Optional[ChannelCollectionResponse]:
        result = await self.app_client.teams.by_team_id(
            team_id=teamID
        ).all_channels.get()

        return result

    async def get_all_messages_in_channel(
        self,
        team_id: str,
        channel_id: str,
    ) -> Optional[ChatMessageCollectionResponse]:
        query_params = (
            GetAllMessagesRequestBuilder.GetAllMessagesRequestBuilderGetQueryParameters(
                filter="lastModifiedDateTime gt 2019-11-01T00:00:00Z"
            )
        )

        #  filter="lastModifiedDateTime gt 2019-11-01T00:00:00Z and lastModifiedDateTime lt 2021-11-01T00:00:00Z",

        request_configuration = RequestConfiguration(
            query_parameters=query_params,
        )

        result = (
            await self.app_client.teams.by_team_id(team_id=team_id)
            .channels.by_channel_id(channel_id=channel_id)
            .messages.get(request_configuration=request_configuration)
        )

        return result
