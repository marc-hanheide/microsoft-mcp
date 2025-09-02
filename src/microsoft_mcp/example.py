import asyncio
from azure.identity import InteractiveBrowserCredential
from msgraph import GraphServiceClient
from os import getenv
from pprint import pprint

credential = InteractiveBrowserCredential(
    client_id=getenv('CLIENT_ID'),
    tenant_id=getenv('TENANT_ID'),
)
scopes = [
    "User.Read",
    "User.ReadBasic.All",
    "Mail.Read",
    "Team.ReadBasic.All",
    "TeamMember.ReadWrite.All"
    ]

client = GraphServiceClient(credentials=credential, scopes=scopes,)





# GET /me
async def me():
    me = await client.me.get()
    #print(await client.teams.get())
    if me:
        print(me)

# Search for emails
async def search_emails(search_query):
    """
    Search for emails using Microsoft Graph SDK

    Args:
        search_query (str): The search query to find emails
    """



    try:
        # Search messages in the user's mailbox
        from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder

        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
    		search = search_query,
        )

        request_configuration = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        messages = await client.me.messages.get(
            request_configuration = request_configuration
        )

        if messages and messages.value:
            print(f"\nFound {len(messages.value)} email(s) matching '{search_query}':")
            print("-" * 80)

            for message in messages.value:
                pprint(f"{message}")
                print("-" * 80)
        else:
            print(f"No emails found matching '{search_query}'")

    except Exception as e:
        print(f"Error searching emails: {e}")

async def search_unified(search_query):
    """
    Search for items using Microsoft Graph SDK

    Args:
        search_query (str): The search query to find items
    """



    try:
        from msgraph.generated.search.query.query_post_request_body import QueryPostRequestBody
        from msgraph.generated.models.search_request import SearchRequest
        from msgraph.generated.models.entity_type import EntityType
        from msgraph.generated.models.search_query import SearchQuery

        request_body = QueryPostRequestBody(
            requests = [
                SearchRequest(
                    entity_types = [
                        EntityType.Message
                    ],
                    query = SearchQuery(
                        query_string = search_query,
                    ),

                    from_ = 0,
                    size = 10,
                    enable_top_results = True,
                ),
            ],
        )

        results = await client.search.query.post(request_body)


        if results and results.value:
            print(f"\nFound {len(results.value)} items matching '{search_query}':")

            for result in results.value:
                #pprint(f"{result}")
                for container in result.hits_containers:
                    for hit in container.hits:
                        print(f"Rank: {hit.rank}")
                        print(f"Summary: {hit.summary}")
                        print(f"Subject: {hit.resource.subject}")
                        print(f"Body: {hit.resource.body.content if hit.resource.body else 'No body content'}")
                        pprint(f"Resource: {hit.resource.__dict__}")
                        print("-" * 80)
        else:
            print(f"No itmes found matching '{search_query}'")

    except Exception as e:
        print(f"Error searching items: {e}")
        raise e

# Main execution
async def main():
    # Get user info
    await me()

    # Search for emails - you can modify the search query as needed
    await search_unified("L-CAS")  # Example: search for emails containing "meeting"

    # You can add more search queries here
    # await search_emails("project")
    # await search_emails("from:example@domain.com")

asyncio.run(main())