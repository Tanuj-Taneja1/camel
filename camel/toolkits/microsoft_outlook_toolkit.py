# ========= Copyright 2023-2024 @ CAMEL-AI.org. All Rights Reserved. =========
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
# ========= Copyright 2023-2024 @ CAMEL-AI.org. All Rights Reserved. =========
import os
from typing import List, Optional

from camel.logger import get_logger
from camel.toolkits import FunctionTool
from camel.toolkits.base import BaseToolkit
from camel.utils import MCPServer, api_keys_required

logger = get_logger(__name__)

local = True

if local:
    from dotenv import load_dotenv

    load_dotenv()


@MCPServer()
class OutlookToolkit(BaseToolkit):
    """A toolkit for interacting with Microsoft Outlook via Graph API."""

    def __init__(
        self,
        timeout: Optional[float] = None,
        scopes: Optional[List[str]] = None,
    ):
        """Initialize the Outlook Toolkit."""
        super().__init__(timeout=timeout)
        if scopes is None:
            scopes = ['https://graph.microsoft.com/.default']
        self.scopes = scopes
        self.credentials = self._authenticate()
        self.client = self._get_graph_client(
            credentials=self.credentials, scopes=self.scopes
        )

    @api_keys_required(
        [
            (None, "TENANT_ID"),
            (None, "MICROSOFT_CLIENT_ID"),
            (None, "MICROSOFT_CLIENT_SECRET"),
        ]
    )
    def _authenticate(self):
        r"""Gets Credentials from environment variables."""
        self.tenant_id = os.getenv("TENANT_ID")
        self.client_id = os.getenv("MICROSOFT_CLIENT_ID")
        self.client_secret = os.getenv("MICROSOFT_CLIENT_SECRET")

        from azure.identity.aio import ClientSecretCredential

        credentials = ClientSecretCredential(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret,
        )
        return credentials

    def _get_graph_client(self, credentials, scopes):
        """Create a client for Microsoft Graph API."""

        from msgraph import GraphServiceClient

        client = GraphServiceClient(credentials=credentials, scopes=scopes)
        return client

    def get_tools(self) -> List[FunctionTool]:
        """Returns a list of FunctionTool objects representing the
        functions in the toolkit.

        """
        return []
