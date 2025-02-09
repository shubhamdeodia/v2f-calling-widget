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
      };
    };
  }
}
export {}; // Necessary when using "declare global" in a module file.

import React, { useEffect, useState } from "react";
import { CallClient, CallAgent, Call } from "@azure/communication-calling";
import { AzureCommunicationTokenCredential } from "@azure/communication-common";
import { CommunicationIdentityClient } from "@azure/communication-identity";

// The ACS connection string is expected to be provided in your environment.
const ACS_CONNECTION_STRING = "endpoint=https://acs-v2f-poc-eastus.unitedstates.communication.azure.com/;accesskey=CJStyM5oFKUlJ9wVkS9d1TmcsPLwSU3GylBu6Elb3hxJEU1WoMXFJQQJ99BBACULyCpQpiE2AAAAAZCSgEAi";

//
// Custom parameters interface. In this implementation, we only expect an ACS user identifier.
// (No default CIF values will be provided.)
//
interface CustomParams {
  acsUser: string;
}

//
// Helper function to generate a token using your ACS connection string.
// If acsUser is provided, that token will be generated for the existing user; otherwise, a new user is created.
async function getTokenUsingConnectionString(
  connectionString: string,
  acsUser?: string
): Promise<{ token: string; userId: string }> {
  const identityClient = new CommunicationIdentityClient(connectionString);
  let user;
  if (acsUser && acsUser.trim() !== "") {
    user = { communicationUserId: acsUser };
  } else {
    // Create a new ACS user if none is provided.
    user = await identityClient.createUser();
  }
  const tokenResponse = await identityClient.getToken(user, ["voip"]);
  return { token: tokenResponse.token, userId: user.communicationUserId };
}

//
// The ACSVoiceWidget component
//
const ACSVoiceWidget: React.FC = () => {
  // Local state variables
  const [phoneState, setPhoneState] = useState<
    "Idle" | "Dialing" | "Ongoing" | "CallSummary"
  >("Idle");
  const [callee, setCallee] = useState<string>("");
  const [callAgent, setCallAgent] = useState<CallAgent | null>(null);
  const [currentCall, setCurrentCall] = useState<Call | null>(null);
  const [callDuration, setCallDuration] = useState<number>(0);
  const [customParams, setCustomParams] = useState<CustomParams | null>(null);

  // Function to initialize ACS using the connection string.
  const initializeACS = async (params: CustomParams) => {
    console.log("[initializeACS] Starting initialization with params:", params);
    try {
      if (!ACS_CONNECTION_STRING) {
        throw new Error("ACS connection string not provided in environment.");
      }
      // Generate a token using the connection string. If params.acsUser is empty, a new ACS user is created.
      const { token, userId } = await getTokenUsingConnectionString(
        ACS_CONNECTION_STRING,
        params.acsUser
      );
      console.log("[initializeACS] Generated token for user:", userId);

      // Create a token refresher that regenerates the token via the connection string.
      const tokenCredential = new AzureCommunicationTokenCredential({
        tokenRefresher: async () => {
          console.log("[initializeACS] Token refresher invoked");
          const result = await getTokenUsingConnectionString(
            ACS_CONNECTION_STRING,
            params.acsUser
          );
          return result.token;
        },
        token: token, // Use the initially generated token.
        refreshProactively: false,
      });
      console.log("[initializeACS] TokenCredential created");

      const callClient = new CallClient();
      console.log(
        "[initializeACS] Creating CallAgent with displayName:",
        params.acsUser || "New ACS User"
      );
      const agent = await callClient.createCallAgent(tokenCredential, {
        displayName: params.acsUser || "New ACS User",
      });
      console.log("[initializeACS] CallAgent created:", agent);
      setCallAgent(agent);

      // Request permission for audio devices.
      console.log("[initializeACS] Requesting audio device permissions...");
      const deviceManager = await callClient.getDeviceManager();
      await deviceManager.askDevicePermission({ audio: true, video: false });
      console.log("[initializeACS] Audio permission granted");

      // Register incoming call event handler.
      agent.on("incomingCall", (args) => {
        console.log("[initializeACS] Incoming call event received:", args);
      });
      setPhoneState("Idle");
      console.log(
        "[initializeACS] Initialization complete; widget state set to Idle"
      );
    } catch (error) {
      console.error("[initializeACS] Error initializing ACS:", error);
    }
  };

  // Retrieve CIF custom parameters if available.
  // Do not provide a default—if CIF APIs are not available, we simply log a warning and initialize ACS with an empty acsUser.
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
        console.log("[fetchCifParams] Received environment object:", envObj);
        // Assume the CIF custom parameter is a JSON string that contains an acsUser value.
        // For example: { "acsUser": "8:acs:xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" }
        const params = JSON.parse(envObj.customParams) as CustomParams;
        console.log(
          "[fetchCifParams] Parsed custom parameters from CIF:",
          params
        );
        setCustomParams(params);
        await initializeACS(params);
      } catch (error) {
        console.error("[fetchCifParams] Error fetching CIF parameters:", error);
      }
    } else {
      console.warn("[fetchCifParams] CIF APIs not available.");
      // Do not provide any default values; simply initialize ACS with an empty acsUser.
      const params: CustomParams = { acsUser: "" };
      setCustomParams(params);
      await initializeACS(params);
    }
    console.log("[fetchCifParams] Exiting fetchCifParams");
  };

  // Handler to place an outgoing call using ACS.
  const handlePlaceCall = async (phoneNumberInput?: string) => {
    console.log("[handlePlaceCall] Entering handlePlaceCall");
    if (!callAgent) {
      console.error("[handlePlaceCall] Call agent is not initialized yet.");
      return;
    }
    const targetNumber = phoneNumberInput || callee;
    if (!targetNumber) {
      console.error("[handlePlaceCall] No valid phone number provided.");
      return;
    }
    console.log("[handlePlaceCall] Placing call to:", targetNumber);
    try {
      // For outbound PSTN calls, ACS accepts a phone number identifier.
      const identifier = { phoneNumber: targetNumber };
      const call = callAgent.startCall([identifier]);
      console.log("[handlePlaceCall] Outgoing call started:", call);
      setCurrentCall(call);
      setPhoneState("Dialing");

      // Subscribe to call state changes.
      call.on("stateChanged", () => {
        console.log("[handlePlaceCall] Call state changed:", call.state);
        if (call.state === "Connected") {
          setPhoneState("Ongoing");
        } else if (call.state === "Disconnected") {
          setPhoneState("CallSummary");
        }
      });
    } catch (error) {
      console.error("[handlePlaceCall] Error placing call:", error);
    }
    console.log("[handlePlaceCall] Exiting handlePlaceCall");
  };

  // Handler to hang up/end the call.
  const handleHangup = () => {
    console.log("[handleHangup] Entering handleHangup");
    if (currentCall) {
      currentCall.hangUp();
      setPhoneState("CallSummary");
      console.log("[handleHangup] Call hung up, state set to CallSummary");
    } else {
      console.warn("[handleHangup] No active call to hang up");
    }
    console.log("[handleHangup] Exiting handleHangup");
  };

  // Timer effect: update call duration when the call is ongoing.
  useEffect(() => {
    console.log("[Timer Effect] Phone state changed to:", phoneState);
    let timerId: number;
    if (phoneState === "Ongoing") {
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

  // Initialize CIF integration on component mount.
  useEffect(() => {
    console.log(
      "[ACSVoiceWidget] Component mounted, initializing CIF integration"
    );
    fetchCifParams();
  }, []);

  // Render the UI based on current phone state.
  return (
    <div style={{ padding: "1rem", fontFamily: "Arial, sans-serif" }}>
      <h3>ACS Voice Calling Widget</h3>
      {customParams ? (
        <p>
          Call Agent initialized for {customParams.acsUser || "New ACS User"}
        </p>
      ) : (
        <p>Initializing ACS...</p>
      )}
      {phoneState === "Idle" && (
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
      {phoneState === "Dialing" && (
        <div>
          <p>Dialing {callee}...</p>
        </div>
      )}
      {phoneState === "Ongoing" && (
        <div>
          <p>Call in progress with {callee}.</p>
          <p>Duration: {callDuration} seconds</p>
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
      {phoneState === "CallSummary" && (
        <div>
          <p>Call ended with {callee}.</p>
          <button
            onClick={() => {
              console.log("[UI] Reset button clicked, resetting state to Idle");
              setPhoneState("Idle");
            }}
          >
            Reset
          </button>
        </div>
      )}
    </div>
  );
};

export default ACSVoiceWidget;
