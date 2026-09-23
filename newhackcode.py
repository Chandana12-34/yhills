from langflow.custom import Component
from langflow.io import MessageTextInput, Output
from langflow.schema import Data


class PAMMockTool(Component):

    display_name = "PAM Mock Tool"
    description = "Mock PAM service for PAM golden scenarios."
    icon = "shield"
    name = "PAMMockTool"

    inputs = [
        MessageTextInput(
            name="request",
            display_name="PAM Request",
            info="Request received from the PAM agent.",
            tool_mode=True,
        ),
    ]

    outputs = [
        Output(
            display_name="PAM Result",
            name="pam_result",
            method="run_pam",
        ),
    ]

    def run_pam(self) -> Data:

        request = (self.request or "").lower().strip()

        # ============================================================
        # MOCK PAM DATABASE
        # ============================================================

        accounts = {

            "acc-1001": {
                "account_id": "ACC-1001",
                "owner": "Rahul",
                "environment": "PROD",
                "status": "ACTIVE",
                "permissions": ["READ", "REPORT_VIEW"],
                "recertification": "CURRENT",
                "owner_tagged": True
            },

            "acc-1002": {
                "account_id": "ACC-1002",
                "owner": "Priya",
                "environment": "UAT",
                "status": "ACTIVE",
                "permissions": ["READ", "WRITE"],
                "recertification": "PENDING",
                "owner_tagged": False
            },

            "acc-1003": {
                "account_id": "ACC-1003",
                "owner": "Amit",
                "environment": "DEV",
                "status": "DISABLED",
                "permissions": ["READ"],
                "recertification": "EXPIRED",
                "owner_tagged": True
            },

            "acc-1004": {
                "account_id": "ACC-1004",
                "owner": "Sneha",
                "environment": "PROD",
                "status": "ACTIVE",
                "permissions": ["READ", "WRITE", "ADMIN"],
                "recertification": "CURRENT",
                "owner_tagged": True
            }
        }

        # ============================================================
        # SERVICE FAILURE SCENARIO
        # ============================================================

        if (
            "service down" in request
            or "service unavailable" in request
            or "vault unavailable" in request
            or "pam unavailable" in request
        ):
            return Data(
                value={
                    "status": "SERVICE_UNAVAILABLE",
                    "message": "PAM service is currently unavailable."
                }
            )

        # ============================================================
        # FIND ACCOUNT ID
        # ============================================================

        matched_account = None

        for account_key, account in accounts.items():

            if account_key in request:
                matched_account = account
                break

        # ============================================================
        # ACCOUNT NOT FOUND
        # ============================================================

        if matched_account is None:

            return Data(
                value={
                    "status": "NOT_FOUND",
                    "message": "No matching PAM account was found.",
                    "action": "DO_NOT_INVENT_DATA"
                }
            )

        account = matched_account

        # ============================================================
        # PERMISSION REVIEW
        # ============================================================

        if (
            "permission" in request
            or "permissions" in request
            or "access" in request
        ):

            return Data(
                value={
                    "status": "SUCCESS",
                    "operation": "PERMISSION_REVIEW",
                    "account_id": account["account_id"],
                    "owner": account["owner"],
                    "environment": account["environment"],
                    "account_status": account["status"],
                    "permissions": account["permissions"]
                }
            )

        # ============================================================
        # RECERTIFICATION
        # ============================================================

        if (
            "recertification" in request
            or "recert" in request
        ):

            if not account["owner_tagged"]:

                return Data(
                    value={
                        "status": "MISSING_OWNER",
                        "operation": "RECERTIFICATION_STATUS",
                        "account_id": account["account_id"],
                        "recertification": account["recertification"],
                        "message": (
                            "Recertification item has no assigned owner. "
                            "Automatic processing is not allowed."
                        )
                    }
                )

            return Data(
                value={
                    "status": "SUCCESS",
                    "operation": "RECERTIFICATION_STATUS",
                    "account_id": account["account_id"],
                    "recertification": account["recertification"],
                    "owner": account["owner"]
                }
            )

        # ============================================================
        # CREDENTIAL REQUEST
        # ============================================================

        if (
            "credential" in request
            or "credentials" in request
            or "password" in request
            or "secret" in request
        ):

            return Data(
                value={
                    "status": "APPROVAL_REQUIRED",
                    "operation": "CREDENTIAL_REQUEST",
                    "account_id": account["account_id"],
                    "environment": account["environment"],
                    "message": (
                        "Privileged credential retrieval requires "
                        "explicit entitlement or approval."
                    ),
                    "credential": None
                }
            )

        # ============================================================
        # NORMAL ACCOUNT SEARCH
        # ============================================================

        return Data(
            value={
                "status": "SUCCESS",
                "operation": "ACCOUNT_SEARCH",
                "account_id": account["account_id"],
                "owner": account["owner"],
                "environment": account["environment"],
                "account_status": account["status"]
            }
        )