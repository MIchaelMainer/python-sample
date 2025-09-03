import asyncio
import os
from datetime import datetime

from azure.identity import DeviceCodeCredential
from dotenv import load_dotenv 

from microsoft_agents_m365copilot_core.src._enums import APIVersion
from microsoft_agents_m365copilot_core.src.client_factory import  MicrosoftAgentsM365CopilotClientFactory

def main():
    load_dotenv()

    TENANT_ID = os.getenv("TENANT_ID")
    CLIENT_ID = os.getenv("CLIENT_ID")

    # Define a proper callback function that accepts all three parameters
    def auth_callback(verification_uri: str, user_code: str, expires_on: datetime):
        print(f"\nTo sign in, use a web browser to open the page {verification_uri}")
        print(f"Enter the code {user_code} to authenticate.")
        print(f"The code will expire at {expires_on}")

    # Create device code credential with correct callback
    credentials = DeviceCodeCredential(
        client_id=CLIENT_ID,
        tenant_id=TENANT_ID,
        prompt_callback=auth_callback
    )

    client = MicrosoftAgentsM365CopilotClientFactory.create_with_default_middleware(api_version=APIVersion.beta)

    client.base_url = "https://graph.microsoft.com/beta" # Make sure the base URL is set to beta

    async def retrieve():
        try:
            # Kick off device code flow and get the token.
            loop = asyncio.get_running_loop()
            token = await loop.run_in_executor(None, lambda: credentials.get_token("https://graph.microsoft.com/.default"))

            # Set the access token.
            headers = {"Authorization": f"Bearer {token.token}"}

            # Print the URL being used.
            print(f"Using API base URL for incoming request: {client.base_url}\n")

            # Directly use httpx to test the endpoint.
            response = await client.post("https://graph.microsoft.com/beta/copilot/retrieval", json={
                "queryString": "What is the latest in my organization?",
                "dataSource": "sharePoint",
                "resourceMetadata": [
                    "title",
                    "author"
                ],
                "maximumNumberOfResults": "10"
            }, headers=headers)

            # Show the response
            print(f"Response HTTP status: {response.status_code}")
            print(f"Response JSON content: {response.text}")
                
        finally:
            print("Your call to the Copilot APIs is now complete.")

    # Run the async function
    asyncio.run(retrieve())

if __name__ == "__main__":
    main()