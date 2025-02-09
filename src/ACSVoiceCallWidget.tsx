// src/ACSVoiceWidget.tsx

// Extend the global Window interface to include the Microsoft object for CIF.
declare global {
  interface Window {
    Microsoft?: {
      CIFramework?: {
        getEnvironment: () => Promise<{ customParams: string }>;
        addHandler: (
          eventName: string,
          handler: (param: string) => void
        ) => void;
        setMode?: (mode: number) => Promise<unknown>;
        setClickToAct?: (value: boolean) => Promise<unknown>;
      };
    };
  }
}
export {}; // Ensure module

import React, { useEffect, useState } from "react";
import { CallClient, CallAgent, Call } from "@azure/communication-calling";
import { AzureCommunicationTokenCredential } from "@azure/communication-common";
import { CommunicationIdentityClient } from "@azure/communication-identity";

// Get ACS connection string from environment variable.
const ACS_CONNECTION_STRING =
  "endpoint=https://acs-v2f-poc-eastus.unitedstates.communication.azure.com/;accesskey=CJStyM5oFKUlJ9wVkS9d1TmcsPLwSU3GylBu6Elb3hxJEU1WoMXFJQQJ99BBACULyCpQpiE2AAAAAZCSgEAi";

/**
 * Helper function to generate an ACS access token.
 * If an existing userId is provided, the token is refreshed; otherwise, a new user is created.
 */
async function getTokenAndIdentity(
  existingUserId?: string
): Promise<{ token: string; userId: string }> {
  if (!ACS_CONNECTION_STRING) {
    throw new Error("ACS connection string not provided in environment.");
  }
  const identityClient = new CommunicationIdentityClient(ACS_CONNECTION_STRING);
  let user;
  if (existingUserId && existingUserId.trim() !== "") {
    user = { communicationUserId: existingUserId };
    const tokenResponse = await identityClient.getToken(user, ["voip"]);
    console.log(
      `[getTokenAndIdentity] Refreshed token for user: ${existingUserId}`
    );
    return { token: tokenResponse.token, userId: existingUserId };
  } else {
    const identityTokenResponse = await identityClient.createUserAndToken([
      "voip",
    ]);
    console.log(
      `[getTokenAndIdentity] Created new identity with ID: ${identityTokenResponse.user.communicationUserId}`
    );
    return {
      token: identityTokenResponse.token,
      userId: identityTokenResponse.user.communicationUserId,
    };
  }
}

/**
 * Custom parameters interface.
 * In our case, we only need the ACS user identifier.
 */
interface CustomParams {
  acsUser: string;
}

/**
 * Enumeration for phone widget state.
 */
enum PhoneWidgetState {
  Idle = "Idle",
  Dialing = "Dialing",
  Ongoing = "Ongoing",
  CallSummary = "CallSummary",
  Incoming = "Incoming",
  CallAccepted = "CallAccepted",
}

/**
 * ACSVoiceWidget component.
 */
const ACSVoiceWidget: React.FC = () => {
  // Local state variables.
  const [phoneState, setPhoneState] = useState<PhoneWidgetState>(
    PhoneWidgetState.Idle
  );
  const [callee, setCallee] = useState<string>("");
  const [callAgent, setCallAgent] = useState<CallAgent | null>(null);
  const [currentCall, setCurrentCall] = useState<Call | null>(null);
  const [callDuration, setCallDuration] = useState<number>(0);
  const [customParams, setCustomParams] = useState<CustomParams | null>(null);
  const [cifEnv, setCifEnv] = useState<{ customParams: string } | null>(null);

  // ------------------- ACS Initialization (Using Connection String) -------------------
  // This effect runs on mount and initializes ACS independently.
  useEffect(() => {
    console.log("[ACSVoiceWidget] Initializing ACS (connection string based)");
    (async () => {
      try {
        // Always create a new identity by passing an empty string.
        const { token, userId } = await getTokenAndIdentity("");
        console.log("[initializeACS] Token generated for user:", userId);
        // Update our custom parameters for display.
        setCustomParams({ acsUser: userId });

        // Create token credential with a token refresher.
        const tokenCredential = new AzureCommunicationTokenCredential({
          tokenRefresher: async () => {
            console.log("[initializeACS] Token refresher invoked");
            const result = await getTokenAndIdentity(userId);
            return result.token;
          },
          token: token,
          refreshProactively: false,
        });
        console.log("[initializeACS] TokenCredential created");

        const callClient = new CallClient();
        console.log(
          "[initializeACS] Creating CallAgent with displayName:",
          userId
        );
        const agent = await callClient.createCallAgent(tokenCredential, {
          displayName: userId,
        });
        console.log("[initializeACS] CallAgent created:", agent);
        setCallAgent(agent);

        // Request permission for audio devices.
        console.log("[initializeACS] Requesting audio device permissions...");
        const deviceManager = await callClient.getDeviceManager();
        await deviceManager.askDevicePermission({ audio: true, video: false });
        console.log("[initializeACS] Audio permission granted");

        // Register incoming call handler.
        agent.on("incomingCall", (args) => {
          console.log("[initializeACS] Incoming call event received:", args);
        });
        setPhoneState(PhoneWidgetState.Idle);
        console.log(
          "[initializeACS] ACS initialization complete; widget state set to Idle"
        );
      } catch (error) {
        console.error(
          "[initializeACS] Error during ACS initialization:",
          error
        );
      }
    })();
  }, []);

  // ------------------- CIF Integration -------------------
  // This effect runs on mount and handles CIF integration.
  useEffect(() => {
    console.log("[ACSVoiceWidget] Initializing CIF integration");
    fetchCifParams();
    registerCifHandlers();
  }, []);

  // fetchCifParams fetches CIF environment and updates state for display.
  const fetchCifParams = async () => {
    console.log("[fetchCifParams] Entering fetchCifParams");
    if (
      window.Microsoft &&
      window.Microsoft.CIFramework &&
      window.Microsoft.CIFramework.getEnvironment
    ) {
      try {
        console.log(
          "[fetchCifParams] CIF API available. Calling getEnvironment()..."
        );
        const envObj = await window.Microsoft.CIFramework.getEnvironment();
        console.log(
          "[fetchCifParams] Received CIF environment object:",
          envObj
        );
        setCifEnv(envObj);
        try {
          const params = JSON.parse(envObj.customParams) as CustomParams;
          console.log("[fetchCifParams] Parsed CIF custom parameters:", params);
          setCustomParams(params);
        } catch (jsonError) {
          console.error(
            "[fetchCifParams] Error parsing CIF customParams.",
            jsonError
          );
        }
      } catch (error) {
        console.error("[fetchCifParams] Error fetching CIF parameters:", error);
      }
    } else {
      console.warn("[fetchCifParams] CIF APIs not available.");
    }
    console.log("[fetchCifParams] Exiting fetchCifParams");
  };

  // registerCifHandlers registers CIF event handlers.
  const registerCifHandlers = () => {
    console.log("[registerCifHandlers] Entering registerCifHandlers");
    if (
      window.Microsoft &&
      window.Microsoft.CIFramework &&
      window.Microsoft.CIFramework.addHandler
    ) {
      console.log("[registerCifHandlers] CIF APIs are available");
      window.Microsoft.CIFramework.addHandler(
        "onclicktoact",
        async (paramStr: string) => {
          console.log(
            "[CIF Handler] onclicktoact event triggered with paramStr:",
            paramStr
          );
          try {
            const params = JSON.parse(paramStr);
            console.log("[CIF Handler] Parsed CIF params:", params);
            const phNo = params.value;
            console.log("[CIF Handler] Parsed phone number:", phNo);
            setCallee(phNo);
            console.log(
              "[CIF Handler] Calling handlePlaceCall with phone:",
              phNo
            );
            await handlePlaceCall(phNo);
          } catch (error) {
            console.error(
              "[CIF Handler] Error in onclicktoact handler:",
              error
            );
          }
        }
      );
      window.Microsoft.CIFramework.addHandler(
        "onmodechanged",
        async (paramStr: string) => {
          console.log("[CIF Handler] onmodechanged event triggered:", paramStr);
        }
      );
      window.Microsoft.CIFramework.addHandler(
        "onpagenavigate",
        async (paramStr: string) => {
          console.log(
            "[CIF Handler] onpagenavigate event triggered:",
            paramStr
          );
        }
      );
      console.log("[registerCifHandlers] CIF event handlers registered");
    } else {
      console.warn(
        "[registerCifHandlers] CIF APIs not available to register handlers"
      );
    }
    console.log("[registerCifHandlers] Exiting registerCifHandlers");
  };

  // ------------------- Call Handling -------------------

  // Places an outgoing call using the CallAgent.
  const handlePlaceCall = async (phoneNumberInput?: string) => {
    console.log("[handlePlaceCall] Entering handlePlaceCall");
    if (!callAgent) {
      console.error("[handlePlaceCall] Call agent is not initialized yet.");
      return;
    }
    // IMPORTANT: Ensure the phone number is in E.164 format (e.g. "+14387730423").
    const targetNumber = phoneNumberInput || callee;
    if (!targetNumber) {
      console.error("[handlePlaceCall] No valid phone number provided.");
      return;
    }
    console.log("[handlePlaceCall] Placing call to:", targetNumber);
    try {
      const identifier = { phoneNumber: targetNumber };
      console.log(
        "[handlePlaceCall] Calling startCall with identifier:",
        identifier
      );
      const call = callAgent.startCall([identifier]);
      console.log("[handlePlaceCall] Outgoing call started:", call);
      setCurrentCall(call);
      setPhoneState(PhoneWidgetState.Dialing);

      // Subscribe to call state changes.
      call.on("stateChanged", () => {
        console.log("[handlePlaceCall] Call state changed:", call.state);
        if (call.state === "Connected") {
          setPhoneState(PhoneWidgetState.Ongoing);
        } else if (call.state === "Disconnected") {
          setPhoneState(PhoneWidgetState.CallSummary);
        }
      });
    } catch (error) {
      console.error("[handlePlaceCall] Error placing call:", error);
    }
    console.log("[handlePlaceCall] Exiting handlePlaceCall");
  };

  // Hangs up the current call.
  const handleHangup = () => {
    console.log("[handleHangup] Entering handleHangup");
    if (currentCall) {
      console.log("[handleHangup] Hanging up call:", currentCall);
      currentCall.hangUp();
      setPhoneState(PhoneWidgetState.CallSummary);
      console.log("[handleHangup] Call hung up, state set to CallSummary");
    } else {
      console.warn("[handleHangup] No active call to hang up");
    }
    console.log("[handleHangup] Exiting handleHangup");
  };

  // ------------------- Timer Effect -------------------
  // Update call duration when call is ongoing.
  useEffect(() => {
    console.log("[Timer Effect] Phone state changed to:", phoneState);
    let timerId: number;
    if (phoneState === PhoneWidgetState.Ongoing) {
      setCallDuration(0);
      timerId = window.setInterval(() => {
        setCallDuration((prev) => prev + 1);
        console.log("[Timer Effect] Updated call duration:", callDuration + 1);
      }, 1000);
    }
    return () => {
      if (timerId) {
        console.log("[Timer Effect] Clearing call duration timer");
        window.clearInterval(timerId);
      }
    };
  }, [phoneState]);

  // ------------------- Render UI -------------------
  return (
    <div style={{ padding: "1rem", fontFamily: "Arial, sans-serif" }}>
      <h3>ACS Voice Calling Widget</h3>
      {customParams ? (
        <p>Call Agent initialized for {customParams.acsUser}</p>
      ) : (
        <p>Initializing ACS...</p>
      )}
      {phoneState === PhoneWidgetState.Idle && (
        <div>
          <label htmlFor="callee">Enter Callee Number:</label>
          <input
            id="callee"
            type="text"
            value={callee}
            onChange={(e) => {
              console.log("[UI] Callee number changed:", e.target.value);
              setCallee(e.target.value);
            }}
            placeholder="+11234567890"
            style={{ marginLeft: "0.5rem", padding: "0.2rem" }}
          />
          <button
            onClick={() => {
              console.log("[UI] Start Call button clicked");
              handlePlaceCall();
            }}
            disabled={!callAgent}
            style={{ marginLeft: "1rem" }}
          >
            Start Call
          </button>
        </div>
      )}
      {phoneState === PhoneWidgetState.Dialing && (
        <div>
          <p>Dialing {callee}...</p>
        </div>
      )}
      {phoneState === PhoneWidgetState.Ongoing && (
        <div>
          <p>Call in progress with {callee}.</p>
          <p>
            Duration: {Math.floor(callDuration / 3600)}:
            {Math.floor((callDuration % 3600) / 60)}:{callDuration % 60}
          </p>
          <button
            onClick={() => {
              console.log("[UI] Hang Up button clicked");
              handleHangup();
            }}
          >
            Hang Up
          </button>
        </div>
      )}
      {phoneState === PhoneWidgetState.CallSummary && (
        <div>
          <p>Call ended with {callee}.</p>
          <button
            onClick={() => {
              console.log("[UI] Reset button clicked, resetting state to Idle");
              setPhoneState(PhoneWidgetState.Idle);
            }}
          >
            Reset
          </button>
        </div>
      )}
      {cifEnv && (
        <div style={{ marginTop: "1rem", fontSize: "0.8rem", color: "#555" }}>
          <strong>CIF Environment:</strong> {JSON.stringify(cifEnv)}
        </div>
      )}
    </div>
  );
};

export default ACSVoiceWidget;
