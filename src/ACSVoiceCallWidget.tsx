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

// Define our phone widget states
enum PhoneWidgetState {
  Idle = "Idle", // No active call
  Dialing = "Dialing", // Outbound call in progress
  Ongoing = "Ongoing", // Call connected
  CallSummary = "CallSummary", // Call ended (post‑call)
  Incoming = "Incoming",
  CallAccepted = "CallAccepted",
}

// Define the structure of our custom parameters passed via CIF
interface CustomParams {
  acsToken: string;
  acsUser: string;
}

const ACSVoiceWidget: React.FC = () => {
  // Local state variables
  const [phoneState, setPhoneState] = useState<PhoneWidgetState>(
    PhoneWidgetState.Idle
  );
  const [callee, setCallee] = useState<string>("");
  const [callAgent, setCallAgent] = useState<CallAgent | null>(null);
  const [currentCall, setCurrentCall] = useState<Call | null>(null);
  const [callDuration, setCallDuration] = useState<number>(0);
  const [customParams, setCustomParams] = useState<CustomParams | null>(null);
  const [cifEnv, setCifEnv] = useState<{ customParams: string } | null>(null);

  // Retrieve CIF custom parameters (including your ACS token and acsUser)
  const fetchCifParams = async () => {
    if (
      window.Microsoft &&
      window.Microsoft.CIFramework &&
      window.Microsoft.CIFramework.getEnvironment
    ) {
      try {
        const envObj = await window.Microsoft.CIFramework.getEnvironment();
        setCifEnv(envObj); // For debugging or further usage.
        // Assume the CIF custom parameter is a JSON string.
        const params = JSON.parse(envObj.customParams) as CustomParams;
        setCustomParams({
          acsToken:
            "eyJhbGciOiJSUzI1NiIsImtpZCI6IjU3Qjg2NEUwQjM0QUQ0RDQyRTM3OTRBRTAyNTAwRDVBNTE5MjA1RjUiLCJ4NXQiOiJWN2hrNExOSzFOUXVONVN1QWxBTldsR1NCZlUiLCJ0eXAiOiJKV1QifQ.eyJza3lwZWlkIjoiYWNzOjg1ZjVjMGMxLWM1NjUtNDBiMS04ODYyLTdmNGRhNGFkNGNlOV8wMDAwMDAyNS04N2ZmLTIzNjMtOWZmYi05YzNhMGQwMGNmYjQiLCJzY3AiOjE3OTIsImNzaSI6IjE3MzkwNjEwMzQiLCJleHAiOjE3MzkxNDc0MzQsInJnbiI6ImFtZXIiLCJhY3NTY29wZSI6InZvaXAiLCJyZXNvdXJjZUlkIjoiODVmNWMwYzEtYzU2NS00MGIxLTg4NjItN2Y0ZGE0YWQ0Y2U5IiwicmVzb3VyY2VMb2NhdGlvbiI6InVuaXRlZHN0YXRlcyIsImlhdCI6MTczOTA2MTAzNH0.BFwuqvJJ5ZtVFwYT253BDZCkTZx2-D95jlSu-fL6fCPGywt04xuD0OfdgQQKdvbSCNxXoHDOim_ypJymwLjFuY31XlSkTC9AmyT7hkV-PWI_N4tK67pv4mTCOSJUZF4ZntzshziPY9y8ZE8B13GPpG94uOa5PcHFrvgl7q-aa3PhkasEzN0tSGRy1CXokI3efZY-gQiVYZuQS-qFw6QJJY6H0wgf2Ad1Q2lmdoTfLR1hma43KOaZeR484Qmd21N4EARD0GEn_Wz5gsXEBUzSyKfEoCU0wgcO0M8E7E0fATcOg02Itjx5c430IZUg-uzAEhFMKcV4LI2ZuqxegRXJwQ",
          acsUser:
            "8:acs:85f5c0c1-c565-40b1-8862-7f4da4ad4ce9_00000025-87ff-2363-9ffb-9c3a0d00cfb4",
        });
        console.log("Retrieved CIF custom params:", params);
        await initializeACS(params);
      } catch (error) {
        console.error("Error fetching CIF parameters:", error);
      }
    } else {
      console.warn("Microsoft.CIFramework is not available.");
    }
  };

  // Initialize ACS CallAgent using the token from custom parameters
  const initializeACS = async (params: CustomParams) => {
    try {
      const tokenCredential = new AzureCommunicationTokenCredential(
        params.acsToken
      );
      const callClient = new CallClient();
      // Optionally pass a displayName (acsUser) during creation
      const agent = await callClient.createCallAgent(tokenCredential, {
        displayName: params.acsUser,
      });
      setCallAgent(agent);
      console.log("ACS CallAgent initialized:", agent);
      // Request permission for audio devices.
      const deviceManager = await callClient.getDeviceManager();
      await deviceManager.askDevicePermission({ audio: true, video: false });
      // Register incoming call event (customize as needed)
      agent.on("incomingCall", (args) => {
        console.log("Incoming call received:", args);
        // You might want to update state or auto-accept the call here.
      });
      // Set initial widget state
      setPhoneState(PhoneWidgetState.Idle);
    } catch (error) {
      console.error("Error initializing ACS:", error);
    }
  };

  // Register CIF event handlers (e.g. click-to-act, mode change, page navigation)
  const registerCifHandlers = () => {
    if (
      window.Microsoft &&
      window.Microsoft.CIFramework &&
      window.Microsoft.CIFramework.addHandler
    ) {
      // When Dynamics passes a phone number, place a call.
      window.Microsoft.CIFramework.addHandler(
        "onclicktoact",
        async (paramStr: string) => {
          try {
            const params = JSON.parse(paramStr);
            const phNo = params.value;
            setCallee(phNo);
            await handlePlaceCall(phNo);
          } catch (error) {
            console.error("Error in onclicktoact handler:", error);
          }
        }
      );
      // Optionally, add other CIF handlers such as "onmodechanged" or "onpagenavigate".
      window.Microsoft.CIFramework.addHandler(
        "onmodechanged",
        async (paramStr: string) => {
          console.log("Panel mode changed:", paramStr);
          // Update UI or state as needed.
        }
      );
      window.Microsoft.CIFramework.addHandler(
        "onpagenavigate",
        async (paramStr: string) => {
          console.log("Page navigation event:", paramStr);
          // Handle page navigation if required.
        }
      );
    }
  };

  // Handler to place an outgoing call using ACS
  const handlePlaceCall = async (phoneNumberInput?: string) => {
    if (!callAgent) {
      console.error("Call agent is not initialized yet.");
      return;
    }
    const targetNumber = phoneNumberInput || callee;
    if (!targetNumber) {
      console.error("Please provide a valid phone number.");
      return;
    }
    try {
      // For an outbound PSTN call, ACS accepts a phone number identifier.
      const identifier = { phoneNumber: targetNumber };
      const call = callAgent.startCall([identifier]);
      setCurrentCall(call);
      setPhoneState(PhoneWidgetState.Dialing);
      console.log("Outgoing call started:", call);
      // Subscribe to call state changes
      call.on("stateChanged", () => {
        console.log("ACS call state changed:", call.state);
        if (call.state === "Connected") {
          setPhoneState(PhoneWidgetState.Ongoing);
        } else if (call.state === "Disconnected") {
          setPhoneState(PhoneWidgetState.CallSummary);
        }
      });
    } catch (error) {
      console.error("Error placing call:", error);
    }
  };

  // Handler to hang up/end the call
  const handleHangup = () => {
    if (currentCall) {
      currentCall.hangUp();
      setPhoneState(PhoneWidgetState.CallSummary);
    }
  };

  // Timer effect: when call is ongoing, update call duration every second
  useEffect(() => {
    let timerId: number;
    if (phoneState === PhoneWidgetState.Ongoing) {
      setCallDuration(0);
      timerId = window.setInterval(() => {
        setCallDuration((prev) => prev + 1);
      }, 1000);
    }
    return () => {
      if (timerId) window.clearInterval(timerId);
    };
  }, [phoneState]);

  // Initialize CIF integration on component mount
  useEffect(() => {
    fetchCifParams();
    registerCifHandlers();
  }, []);

  // Render UI based on current phone state
  return (
    <div style={{ padding: "1rem", fontFamily: "Arial, sans-serif" }}>
      <h3>ACS Voice Calling Widget</h3>
      {customParams ? (
        <p>Call Agent initialized for {customParams.acsUser}</p>
      ) : (
        <p>Waiting for CIF custom parameters...</p>
      )}
      {phoneState === PhoneWidgetState.Idle && (
        <div>
          <label htmlFor="callee">Enter Callee Number:</label>
          <input
            id="callee"
            type="text"
            value={callee}
            onChange={(e) => setCallee(e.target.value)}
            placeholder="+11234567890"
            style={{ marginLeft: "0.5rem", padding: "0.2rem" }}
          />
          <button
            onClick={() => handlePlaceCall()}
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
          <p>Call with {callee} is in progress.</p>
          <p>
            Duration: {Math.floor(callDuration / 3600)}:
            {Math.floor((callDuration % 3600) / 60)}:{callDuration % 60}
          </p>
          <button onClick={handleHangup}>Hang Up</button>
        </div>
      )}
      {phoneState === PhoneWidgetState.CallSummary && (
        <div>
          <p>Call ended with {callee}.</p>
          <button onClick={() => setPhoneState(PhoneWidgetState.Idle)}>
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
