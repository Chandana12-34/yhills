from langflow.custom import Component
from langflow.io import MessageTextInput, Output
from langflow.schema import Data


class PAMMockTool(Component):

    display_name = "PAM Mock Tool"
    description = "Mock PAM service used for PAM golden scenarios."
    icon = "shield"
    name = "PAMMockTool"

    inputs = [
        MessageTextInput(
            name="request",
            display_name="PAM Request",
            info="Request from the PAM agent.",
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

        # Make sure the input exists
        request = (self.request or "").lower().strip()

        # ------------------------------------------------
        # MOCK ACCOUNT DATA
        # ------------------------------------------------

        accounts = {
            "acc-1001": {
                "account_id": "ACC-1001",
                "owner": "Rahul",
                "environment": "PROD",
                "status": "ACTIVE",
                "permissions": [
                    "READ",
                    "REPORT_VIEW"
                ],
                "recertification": "CURRENT",
                "owner_tagged": True
            },

            "acc-1002": {
                "account_id": "ACC-1002",
                "owner": "Priya",
                "environment": "UAT",
                "status": "ACTIVE",
                "permissions": [
                    "READ",
                    "WRITE"
                ],
                "recertification": "PENDING",
                "owner_tagged": False
            }
        }

        # ------------------------------------------------
        # SERVICE-DOWN SIMULATION
        # ------------------------------------------------

        if (
            "service down" in request
            or "service unavailable" in request
            or "vault unavailable" in request
            or "vault down" in request
        ):
            return Data(
                value={
                    "status": "SERVICE_UNAVAILABLE",
                    "message": "PAM Vault service is currently unavailable."
                }
            )

        # ------------------------------------------------
        # ACCOUNT SEARCH
        # ------------------------------------------------

        for account_id, account in accounts.items():

            if account_id in request:

                # ------------------------------------------------
                # PERMISSION REQUEST
                # ------------------------------------------------

                if (
                    "permission" in request
                    or "permissions" in request
                ):
                    return Data(
                        value={
                            "status": "SUCCESS",
                            "operation": "PERMISSION_REVIEW",
                            "account_id": account["account_id"],
                            "owner": account["owner"],
                            "environment": account["environment"],
                            "permissions": account["permissions"],
                            "account_status": account["status"]
                        }
                    )

                # ------------------------------------------------
                # RECERTIFICATION REQUEST
                # ------------------------------------------------

                if (
                    "recert" in request
                    or "recertification" in request
                ):

                    # Missing owner scenario
                    if not account["owner_tagged"]:
                        return Data(
                            value={
                                "status": "MISSING_OWNER",
                                "operation": "RECERTIFICATION_STATUS",
                                "account_id": account["account_id"],
                                "recertification": account["recertification"],
                                "message": (
                                    "No owner is assigned to this "
                                    "recertification item."
                                )
                            }
                        )

                    # Normal recertification
                    return Data(
                        value={
                            "status": "SUCCESS",
                            "operation": "RECERTIFICATION_STATUS",
                            "account_id": account["account_id"],
                            "recertification": account["recertification"],
                            "owner": account["owner"]
                        }
                    )

                # ------------------------------------------------
                # CREDENTIAL / PASSWORD REQUEST
                # ------------------------------------------------

                if (
                    "credential" in request
                    or "credentials" in request
                    or "password" in request
                    or "secret" in request
                ):
                    return Data(
                        value={
                            "status": "APPROVAL_REQUIRED",
                            "operation": "CREDENTIAL_RETRIEVAL",
                            "account_id": account["account_id"],
                            "message": (
                                "Privileged credential retrieval "
                                "requires explicit approval."
                            ),
                            "credential": None
                        }
                    )

                # ------------------------------------------------
                # ENVIRONMENT REQUEST
                # ------------------------------------------------

                if "environment" in request:
                    return Data(
                        value={
                            "status": "SUCCESS",
                            "operation": "ACCOUNT_ENVIRONMENT",
                            "account_id": account["account_id"],
                            "environment": account["environment"]
                        }
                    )

                # ------------------------------------------------
                # OWNER REQUEST
                # ------------------------------------------------

                if "owner" in request:
                    return Data(
                        value={
                            "status": "SUCCESS",
                            "operation": "ACCOUNT_OWNER",
                            "account_id": account["account_id"],
                            "owner": account["owner"]
                        }
                    )

                # ------------------------------------------------
                # NORMAL ACCOUNT SEARCH
                # ------------------------------------------------

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

        # ------------------------------------------------
        # UNKNOWN ACCOUNT
        # ------------------------------------------------

        return Data(
            value={
                "status": "NOT_FOUND",
                "operation": "ACCOUNT_SEARCH",
                "message": "No matching PAM account was found."
            }
        )