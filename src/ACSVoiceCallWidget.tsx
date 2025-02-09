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
        openForm?: (
          entityFormOptions: string,
          formParameters: string
        ) => Promise<void>;
        setMode?: (mode: number) => Promise<unknown>;
        setClickToAct?: (value: boolean) => Promise<unknown>;
      };
    };
  }
}
export {}; // Ensure this file is treated as a module

import React, { useEffect, useState } from "react";
import { CallClient, CallAgent, Call } from "@azure/communication-calling";
import { AzureCommunicationTokenCredential } from "@azure/communication-common";
import { CommunicationIdentityClient } from "@azure/communication-identity";

// Get your ACS connection string and your ACS phone number from environment variables.
const ACS_CONNECTION_STRING =
  "endpoint=https://acs-v2f-poc-eastus.unitedstates.communication.azure.com/;accesskey=CJStyM5oFKUlJ9wVkS9d1TmcsPLwSU3GylBu6Elb3hxJEU1WoMXFJQQJ99BBACULyCpQpiE2AAAAAZCSgEAi";

/**
 * Helper function that uses your ACS connection string to generate a token.
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
 * Custom parameters interface. We only care about the ACS user identifier.
 */
interface CustomParams {
  acsUser: string;
}

/**
 * Enumeration for phone widget states.
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
  const [phoneState, setPhoneState] = useState<PhoneWidgetState>(
    PhoneWidgetState.Idle
  );
  const [callee, setCallee] = useState<string>("");
  const [callAgent, setCallAgent] = useState<CallAgent | null>(null);
  const [currentCall, setCurrentCall] = useState<Call | null>(null);
  const [callDuration, setCallDuration] = useState<number>(0);
  const [customParams, setCustomParams] = useState<CustomParams | null>(null);
  const [cifEnv, setCifEnv] = useState<{ customParams: string } | null>(null);

  // -------------------- ACS Initialization --------------------
  useEffect(() => {
    console.log("[ACSVoiceWidget] Checking for CIF API:", window.Microsoft && window.Microsoft.CIFramework);
    console.log("[ACSVoiceWidget] Initializing ACS (connection string based)");
    (async () => {
      try {
        // Always create a new identity (pass an empty string).
        const { token, userId } = await getTokenAndIdentity("");
        console.log("[initializeACS] Token generated for user:", userId);
        // Update our custom parameters for display.
        setCustomParams({ acsUser: "+14387730423" });

        // Create a token credential with a token refresher.
        const tokenCredential = new AzureCommunicationTokenCredential({
          tokenRefresher: async () => {
            const result = await getTokenAndIdentity(userId);
            return result.token;
          },
          token: token,
          refreshProactively: true,
        });

        console.log("[initializeACS] TokenCredential created");

        const callClient = new CallClient();
        console.log(
          "[initializeACS] Creating CallAgent with displayName:",
          userId
        );
        // IMPORTANT: For PSTN calls, supply a valid ACS phone number as the caller ID.
        const agent = await callClient.createCallAgent(tokenCredential, {
          displayName: "V2F Demo",
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
          // (You could update state here if you want to display an incoming call UI.)
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

  // -------------------- CIF Initialization --------------------
  // Enable click-to-act as soon as the widget loads.
  useEffect(() => {
    if (
      window.Microsoft &&
      window.Microsoft.CIFramework &&
      window.Microsoft.CIFramework.setClickToAct
    ) {
      window.Microsoft.CIFramework.setClickToAct(true)
        .then(() => console.log("[CIF] Click-to-act enabled"))
        .catch((err) =>
          console.error("[CIF] Error enabling click-to-act:", err)
        );
    }
  }, []);

  // fetchCifParams: fetches CIF environment and updates state for display.
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

  // registerCifHandlers: registers CIF event handlers (onclicktoact, onmodechanged, onpagenavigate).
  const registerCifHandlers = () => {
    console.log("[registerCifHandlers] Entering registerCifHandlers");
    if (
      window.Microsoft &&
      window.Microsoft.CIFramework &&
      window.Microsoft.CIFramework.addHandler
    ) {
      console.log("[registerCifHandlers] CIF APIs are available");
      // Handler for click-to-act: for example, when a user clicks a phone number in Dynamics 365.
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
          // You can add logic here to adjust your UI (for example, expanding or collapsing your widget)
        }
      );
      window.Microsoft.CIFramework.addHandler(
        "onpagenavigate",
        async (paramStr: string) => {
          console.log(
            "[CIF Handler] onpagenavigate event triggered:",
            paramStr
          );
          // You might want to use this event to update context (e.g. record IDs) in your widget
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

  // Initialize CIF integration.
  // Wait for the CIF framework to signal that it is ready (via the "CIFInitDone" event).
  useEffect(() => {
    console.log(
      "[ACSVoiceWidget] Waiting for CIFInitDone event to initialize CIF integration."
    );

    const cifInitHandler = () => {
      console.log("[ACSVoiceWidget] CIFInitDone event received.");
      fetchCifParams();
      registerCifHandlers();
    };

    cifInitHandler(); // Call the handler immediately in case the event has already fired.
    // if (window.Microsoft && window.Microsoft.CIFramework) {
    //   window.addEventListener("CIFInitDone", cifInitHandler);
    // } else {
    //   console.warn(
    //     "[ACSVoiceWidget] CIF APIs not available. Skipping CIF integration."
    //   );
    // }

    // Optionally, clean up the event listener when the component unmounts.
    // return () => {
    //   if (window.Microsoft && window.Microsoft.CIFramework) {
    //     window.removeEventListener("CIFInitDone", cifInitHandler);
    //   }
    // };
  }, []);

  // -------------------- Call Handling --------------------
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
      const identifier = targetNumber.startsWith("+")
        ? { phoneNumber: targetNumber } // For PSTN
        : { communicationUserId: targetNumber }; // For ACS users

      // Replace with your ACS phone number for caller ID
      const alternateCallerId = { phoneNumber: "+18883958119" };

      console.log(
        "[handlePlaceCall] Calling startCall with identifier:",
        identifier
      );
      const call = callAgent.startCall([identifier], { alternateCallerId });
      console.log("[handlePlaceCall] Outgoing call started:", call);

      setCurrentCall(call);
      setPhoneState(PhoneWidgetState.Dialing);

      // Subscribe to call state changes.
      call.on("stateChanged", () => {
        console.log("[handlePlaceCall] Call state changed:", call.state);
        if (call.state === "Connected") {
          console.log("Ongoing call", call);
          setPhoneState(PhoneWidgetState.Ongoing);
        } else if (call.state === "Disconnected") {
          console.log("Disconnected call", call);
          setPhoneState(PhoneWidgetState.CallSummary);
          setCurrentCall(null);
        }
      });
    } catch (error) {
      console.error("[handlePlaceCall] Error placing call:", error);
    }
    console.log("[handlePlaceCall] Exiting handlePlaceCall");
  };

  const handleHangup = () => {
    console.log("[handleHangup] Entering handleHangup");
    if (currentCall) {
      console.log("[handleHangup] Hanging up call:", currentCall);
      currentCall.hangUp();
      setCurrentCall(null);
      setPhoneState(PhoneWidgetState.CallSummary);
      console.log("[handleHangup] Call hung up, state set to CallSummary");
    } else {
      console.warn("[handleHangup] No active call to hang up");
    }
    console.log("[handleHangup] Exiting handleHangup");
  };

  // -------------------- Timer Effect --------------------
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

  // -------------------- Render UI --------------------
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
            placeholder="+14387730423"
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
