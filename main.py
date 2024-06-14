import asyncio
import configparser
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from graph import Graph

MARK_8_PROJECT_TEAM_ID = "4cb230df-8541-4e96-9a8b-6dadf271e2b8"
DESIGN_CHANNEL_ID = "19:0f0889e58a384abda07c8274fbdc125b@thread.tacv2"


async def main():
    print("Python Graph App-Only Tutorial\n")

    # Load settings
    config = configparser.ConfigParser()
    config.read(["config.cfg", "config.dev.cfg"])
    azure_settings = config["azure"]

    graph: Graph = Graph(azure_settings)

    choice = -1

    while choice != 0:
        print("Please choose one of the following options:")
        print("0. Exit")
        print("1. Display access token")
        print("2. List users")
        print("3. Get all teams")
        print("4. Get all Channels in Team")
        print("5. Get all Messages in a Channel in a Team")

        try:
            choice = int(input())
        except ValueError:
            choice = -1

        try:
            if choice == 0:
                print("Goodbye...")
            elif choice == 1:
                await display_access_token(graph)
            elif choice == 2:
                await list_users(graph)
            elif choice == 3:
                await get_teams(graph)
            elif choice == 4:
                await get_channels(graph)

            elif choice == 5:
                await get_all_messages_in_channel(graph)
            else:
                print("Invalid choice!\n")
        except ODataError as odata_error:
            print("Error:")
            if odata_error.error:
                print(odata_error.error.code, odata_error.error.message)


async def display_access_token(graph: Graph):
    token = await graph.get_app_only_token()
    print("App-only token:", token, "\n")


async def list_users(graph: Graph):
    users_page = await graph.get_users()

    # Output each users's details
    if users_page and users_page.value:
        for user in users_page.value:
            print("User:", user.display_name)
            print("  ID:", user.id)
            print("  Email:", user.mail)

        # If @odata.nextLink is present
        more_available = users_page.odata_next_link is not None
        print("\nMore users available?", more_available, "\n")


async def make_graph_call(graph: Graph):
    # TODO
    return


async def get_channels(graph: Graph):
    result = await graph.get_all_channels_in_team(teamID=MARK_8_PROJECT_TEAM_ID)

    if result and result.value:
        for channel in result.value:
            print("Channel name:", channel.display_name)
            print("Channel id:", channel.id)
            print("//////")
    return


async def get_teams(graph: Graph):
    result = await graph.get_all_teams()

    if result and result.value:
        for team in result.value:
            print("Team:", team.display_name)
            print("  ID:", team.id)
            print("  Description:", team.description)

        # If @odata.nextLink is present
        more_available = result.odata_next_link is not None
        print("\nMore teams available?", more_available, "\n")


async def get_all_messages_in_channel(graph: Graph):
    result = await graph.get_all_messages_in_channel(
        channel_id=DESIGN_CHANNEL_ID, team_id=MARK_8_PROJECT_TEAM_ID
    )

    if result and result.value:
        for message in result.value:
            print("  Message:", message.body.content)
            print("  Summary:", message.summary)
            print("  Subject:", message.subject)
            print("  Sender:", message.from_.user.display_name)
            print("  ID:", message.id)
            print("  Attachments", message.attachments)
            print("  Additional Data", message.additional_data)

            print("\n////// ------ //////\n")

        # If @odata.nextLink is present
        more_available = result.odata_next_link is not None
        print("\nMore messages available?", more_available, "\n")


# Run main
asyncio.run(main())


"""
[
'additional_data', 'attachments', 'backing_store', 'body', 'channel_identity', 'chat_id', 'create_from_discriminator_value', 
'created_date_time', 'deleted_date_time', 'etag', 'event_detail', 'from_', 'get_field_deserializers', 'hosted_contents',
 'id', 'importance', 'last_edited_date_time', 'last_modified_date_time', 'locale', 'mentions', 'message_history', 
 'message_type', 'odata_type', 'policy_violation', 'reactions', 'replies', 'reply_to_id', 'serialize', 'subject',
   'summary', 'web_url']
"""
