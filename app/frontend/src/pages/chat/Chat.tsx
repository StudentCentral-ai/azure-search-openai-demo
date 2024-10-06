import { useRef, useState, useEffect, useContext } from "react";
import { useTranslation } from "react-i18next";
import { Helmet } from "react-helmet-async";
import { Checkbox, Panel, DefaultButton, TextField, ITextFieldProps, ICheckboxProps } from "@fluentui/react";
import { SparkleFilled } from "@fluentui/react-icons";
import { useId } from "@fluentui/react-hooks";
import readNDJSONStream from "ndjson-readablestream";

import styles from "./Chat.module.css";

import {
    chatApi,
    configApi,
    RetrievalMode,
    ChatAppResponse,
    ChatAppResponseOrError,
    ChatAppRequest,
    ResponseMessage,
    VectorFieldOptions,
    GPT4VInput,
    SpeechConfig
} from "../../api";
import { Answer, AnswerError, AnswerLoading } from "../../components/Answer";
import { QuestionInput } from "../../components/QuestionInput";
import { ExampleList } from "../../components/Example";
import { UserChatMessage } from "../../components/UserChatMessage";
import { HelpCallout } from "../../components/HelpCallout";
import { AnalysisPanel, AnalysisPanelTabs } from "../../components/AnalysisPanel";
import { SettingsButton } from "../../components/SettingsButton";
// import { ClearChatButton } from "../../components/ClearChatButton";
import { UploadFile } from "../../components/UploadFile";
import { useLogin, getToken, requireAccessControl } from "../../authConfig";
import { VectorSettings } from "../../components/VectorSettings";
import { useMsal } from "@azure/msal-react";
import { TokenClaimsDisplay } from "../../components/TokenClaimsDisplay";
import { GPT4VSettings } from "../../components/GPT4VSettings";
import { LoginContext } from "../../loginContext";
import { LanguagePicker } from "../../i18n/LanguagePicker";

// PromptRenderer
import PromptRenderer from "../../components/PromptRenderer/PromptRenderer";
import { createRoot } from "react-dom/client";

// realtime
import { Player } from "../../components/Realtime/player.ts";
import { Recorder } from "../../components/Realtime/recorder.ts";
import { LowLevelRTClient, SessionUpdateMessage } from "rt-client";

// PromptRenderer
// const samplePrompt = `Please answer to the following question : <mcqdata>{"question":"Which one is a vegetable?", "choices": {"A": "apple","B": "orange", "C": "tomato"}}</mcqdata> What is your answer ?`;
const q1 = `<mcqdata>{"question":"Which of the following is considered a risk-free interest rate?", "choices": {"A": "LIBOR","B": "Repo rate", "C": "Treasury rate", "D": "Fed funds rate", "E": "I don’t know yet"}}</mcqdata>`;
const q2 = `<mcqdata>{"question":"Here is the formula of compounded interest to an initial Investment “A” where “R” is the interest rate and “m” is the compounding frequency:
A(1+R/m)mn
True or False: Increasing the compounding frequency “m” (e.g., from annual to monthly or daily compounding) increases the terminal value of the investment because interest is applied more frequently, allowing for interest to be earned on previously accumulated interest.", "choices": {"A": "True","B": "False"}}</mcqdata>
`;

// Realtime

let prmpt: string = "";

class Lock {
    private _locked: boolean = false;
    private _waiting: Array<() => void> = [];

    async acquire(): Promise<void> {
        if (this._locked) {
            await new Promise<void>(resolve => this._waiting.push(resolve));
        }
        this._locked = true;
    }

    release(): void {
        if (this._waiting.length > 0) {
            const resolve = this._waiting.shift();
            if (resolve) {
                resolve();
            }
        } else {
            this._locked = false;
        }
    }
}

const Chat = () => {
    const [isConfigPanelOpen, setIsConfigPanelOpen] = useState(false);
    const [promptTemplate, setPromptTemplate] = useState<string>("");
    const [temperature, setTemperature] = useState<number>(0.3);
    const [seed, setSeed] = useState<number | null>(null);
    const [minimumRerankerScore, setMinimumRerankerScore] = useState<number>(0);
    const [minimumSearchScore, setMinimumSearchScore] = useState<number>(0);
    const [retrieveCount, setRetrieveCount] = useState<number>(3);
    const [retrievalMode, setRetrievalMode] = useState<RetrievalMode>(RetrievalMode.Hybrid);
    const [useSemanticRanker, setUseSemanticRanker] = useState<boolean>(true);
    const [shouldStream, setShouldStream] = useState<boolean>(true);
    const [useSemanticCaptions, setUseSemanticCaptions] = useState<boolean>(false);
    const [excludeCategory, setExcludeCategory] = useState<string>("");
    const [useSuggestFollowupQuestions, setUseSuggestFollowupQuestions] = useState<boolean>(false);
    const [vectorFieldList, setVectorFieldList] = useState<VectorFieldOptions[]>([VectorFieldOptions.Embedding]);
    const [useOidSecurityFilter, setUseOidSecurityFilter] = useState<boolean>(false);
    const [useGroupsSecurityFilter, setUseGroupsSecurityFilter] = useState<boolean>(false);
    const [gpt4vInput, setGPT4VInput] = useState<GPT4VInput>(GPT4VInput.TextAndImages);
    const [useGPT4V, setUseGPT4V] = useState<boolean>(false);

    const lastQuestionRef = useRef<string>("");
    const chatMessageStreamEnd = useRef<HTMLDivElement | null>(null);

    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [isStreaming, setIsStreaming] = useState<boolean>(false);
    const [error, setError] = useState<unknown>();

    const [activeCitation, setActiveCitation] = useState<string>();
    const [activeAnalysisPanelTab, setActiveAnalysisPanelTab] = useState<AnalysisPanelTabs | undefined>(undefined);

    const [selectedAnswer, setSelectedAnswer] = useState<number>(0);
    const [answers, setAnswers] = useState<[user: string, response: ChatAppResponse][]>([]);
    const [streamedAnswers, setStreamedAnswers] = useState<[user: string, response: ChatAppResponse][]>([]);
    const [speechUrls, setSpeechUrls] = useState<(string | null)[]>([]);

    const [showGPT4VOptions, setShowGPT4VOptions] = useState<boolean>(false);
    const [showSemanticRankerOption, setShowSemanticRankerOption] = useState<boolean>(false);
    const [showVectorOption, setShowVectorOption] = useState<boolean>(false);
    const [showUserUpload, setShowUserUpload] = useState<boolean>(false);
    const [showLanguagePicker, setshowLanguagePicker] = useState<boolean>(false);
    const [showSpeechInput, setShowSpeechInput] = useState<boolean>(false);
    const [showSpeechOutputBrowser, setShowSpeechOutputBrowser] = useState<boolean>(false);
    const [showSpeechOutputAzure, setShowSpeechOutputAzure] = useState<boolean>(false);
    const audio = useRef(new Audio()).current;
    const [isPlaying, setIsPlaying] = useState(false);

    const [workflowStateNo, setWorkflowStateNo] = useState<number>(0);

    const [selectedMCQAnswer, setSelectedMCQAnswer] = useState("");

    const handleMCQAnswerSelected = (value: string) => {
        setSelectedMCQAnswer(value);
        console.log("Selected MCQ Answer:", value);
    };

    // PromptRenderer
    const [mcqPrompt, setMcqPrompt] = useState<string>("");
    const [studentTurn, setStudentTurn] = useState<string>("");

    // Realtime
    let realtimeStreaming: LowLevelRTClient;
    let audioRecorder: Recorder;
    let audioPlayer: Player;

    const lock = new Lock();

    const ENDPOINT = "https://stce-aiexp-core-dev-aoai.openai.azure.com/";
    const API_KEY = "OPENAI_API_KEY_SECRET";
    const DEPLOYMENT = "gpt-4o-realtime-preview-global";
    const TEMPERATURE = 0.8;
    const VOICE = "echo";
    const SYSTEM_PROMPT = `System Prompt: You are Greg, an empathetic, knowledgeable and encouraging tutor who assists students in reviewing their coursework and preparing effectively for exams.
You possess academic expertise and teaching skills to engage in discussions on any course topic, guiding the conversation through questions in the style of a Socratic Dialogue. You can propose quantitative exercise and assess the student’s step-by-step reasoning as they progress towards the solution.
You prefer to assist students by asking guiding questions. 
When a student asks a question, you respond with another simple question to help them gradually find the solution on their own. You only provide direct answers when you sense the student is truly stuck and it's more beneficial to move forward.
Respond in a casual and friendly tone.
Sprinkle in filler words, contractions, idioms, and other casual speech that we use in conversation.
Emulate the user’s speaking style while maintaining a warm and supportive tone, like a friendly tutor.
If the user drifts from the topic of the course, gently steer the conversation back to this topic.
Each of your utterances includes a brief comment, followed by either a new question or an encouraging message to motivate the student to continue their response.
Be concise, limiting your utterances to 150 words or less.
It is extremely important that any JSON data shall be rendered exactly as you have memorized it.
If you are representing math formulas, please respect strict rule of wrapping any formula between delimiters: $$...$$ for block math or $...$ for inline math.
The language spoken by the student and you is: {{{language}}}. You will discuss with the student in the requested language, with the native accent from the original country of the anguage.

---

First message: “Hi Nicolas! Great to have you here— I’m looking forward to a great tutoring session together today!”
Start the session by assessing:
•	if the student has sufficient time for the session (30 min to 1h is required)
Then ask,
•	if the student is in an appropriate work setting (quiet space, stable internet connection)

Conclude this check by inviting the student to start the Tutoring session
The Tutoring session is composed of 3 distinct sections: 
1) Quiz 
2) Open-Ended Questions 
3) Quantitative Exercise


QUIZ
Context: MCQ and true/false questions to be displayed on screen for the student to answer. Give the MCQ in the exact form as you have memorized, JSON format including.
Once the student has answered an answer, ask for a comment on why this choice. You prefer to assist students by asking guiding questions. 

QUESTION 1. {"question":"Which of the following is considered a risk-free interest rate?", "choices": {"A": "LIBOR","B": "Repo rate", "C": "Treasury rate", "D": "Fed funds rate", "E": "I don’t know yet"}}

Correct answer: C
Ask the student to elaborate on his selected answer. 
Then, if the student selected the wrong answer, ask the definition of “C. Treasury Rate”, and whether it could be the right answer.

Context: MCQ and true/false questions to be displayed on screen for the student to answer. Give the MCQ in the exact form as you memorize, tags and JSON format including. 
Once the student has answered an answer, ask for a comment on why this choice. You prefer to assist students by asking guiding questions. 

QUESTION 2. {"question":"Here is the formula of compounded interest to an initial Investment “A” where “R” is the interest rate and “m” is the compounding frequency: A(1+R/m)mn. True or False: Increasing the compounding frequency “m” (e.g., from annual to monthly or daily compounding) increases the terminal value of the investment because interest is applied more frequently, allowing for interest to be earned on previously accumulated interest.", "choices": {"A": "True","B": "False"}}

Correct Answer: True
Ask the student to elaborate on his selected answer. 
Once the student has elaborated on his answer, ask him “what is “n” in the formula

Second part of the tutoring session: 2 Open-Ended discussions to be discussed orally
Open-Ended Questions
Discussion 1:
“In the first question of the quiz, LIBOR was mentioned. What do you know about the LIBOR? “
Assess the student’s answer – is it correct? it is complete? Ask for additional comments or clarifications in case the answer can be improved. You prefer to assist students by asking guiding questions. 
Discussion 2:
Imagine you manage a bond portfolio, and interest rates for short-term bonds are usually given with semi-annual compounding. If the central bank changes its policy and the market starts quoting rates using continuous compounding, how would this affect your portfolio management?
This second discussion is harder than the first one. The student will most likely struggle to answer, help him by going step by step and asking simple guiding questions
-	guiding questions on the difference between the 2 compounding methods
-	guiding questions on the equation which allows to convert one rate into another.
If after 4 questions in total, the answer of the student is still incomplete, conclude this discussion by inviting the student to review his course on his own before the next session.

Third part of the tutoring session: Quantitative Exercises
Quantitative Exercises

EXERCISE 1 - When a bank states that the interest rate on one-year deposits is 10% per annum with semi-annual compounding, how much will $100 grow to at the end of one year?  
Correct answer: $110.25
If the student is stuck on this question, ask him to clarify the formula we should apply for each case.
Once this formula has been clarified, let the student solve the equation for each one. 

EXERCISE 2 - When a bank states that the interest rate on one-year deposits is 10% per annum with continuous compounding, how much will $100 grow to at the end of one year?  
Correct answer: $110.52
If the student is stuck on this question, ask him to clarify the formula we should apply for each case.

Conclude this tutoring session by asking the student 
-	when he is available for the next session
-	to review what he has learnt so far, and get ready with the next chapter of his course: “bond pricing”


`;

    /**
     * Starts the real-time audio streaming process.
     *
     * This function initializes the LowLevelRTClient with the provided endpoint, API key,
     * and deployment or model. It then sends a session configuration message to the server.
     * If the configuration message fails to send, it logs an error and updates the UI state.
     *
     * @param endpoint - The endpoint URL for the real-time audio service.
     * @param apiKey - The API key for authenticating with the service.
     * @param deploymentOrModel - The deployment or model identifier for the service.
     *
     * @returns A promise that resolves when the audio streaming process has started and
     *          the initial configuration message has been sent.
     */
    async function start_realtime() {
        const endpoint = ENDPOINT;
        const apiKey = API_KEY;
        const deploymentOrModel = DEPLOYMENT;

        if (!endpoint && !deploymentOrModel) {
            alert("Endpoint and Deployment are required for Azure OpenAI");
            return;
        }

        // if (!deploymentOrModel) {
        //     alert("Model is required for OpenAI");
        //     return;
        // }

        if (!apiKey) {
            alert("API Key is required");
            return;
        }

        console.log("start_realtime: endpoint: " + endpoint + ", apiKey: ***********" + ", deploymentOrModel: " + deploymentOrModel);
        realtimeStreaming = new LowLevelRTClient(new URL(endpoint), { key: apiKey }, { deployment: deploymentOrModel });
        try {
            console.log("start_realtime: sending session config");
            await realtimeStreaming.send(createConfigMessage());
            console.log("start_realtime: sent");
        } catch (error) {
            console.log(error);
            //makeNewTextBlock("[Connection error]: Unable to send initial config message. Please check your endpoint and authentication details.");
            setMcqPrompt("[Connection error]: Unable to send initial config message. Please check your endpoint and authentication details.");
            setFormInputState(InputState.ReadyToStart);
            return;
        }

        try {
            // await Promise.all([resetAudio(true), handleRealtimeMessages()]);
            console.log("start_realtime: resetting audio and handling messages");
            await resetAudio(true);
            console.log("start_realtime: reset audio complete. Handling messages...");
            await handleRealtimeMessages();
            console.log("start_realtime: handling messages complete");
        } catch (error) {
            console.log("start_realtime", error);
        }
    }

    function createConfigMessage(): SessionUpdateMessage {
        let configMessage: SessionUpdateMessage = {
            type: "session.update",
            session: {
                turn_detection: {
                    type: "server_vad"
                },
                input_audio_transcription: {
                    model: "whisper-1"
                }
            }
        };

        const systemMessage = getSystemMessage();
        const temperature = getTemperature();
        const voice = getVoice();

        if (systemMessage) {
            configMessage.session.instructions = systemMessage;
        }
        if (!isNaN(temperature)) {
            configMessage.session.temperature = temperature;
        }
        if (voice) {
            configMessage.session.voice = voice;
        }
        console.log("configMessage: " + JSON.stringify(configMessage));
        return configMessage;
    }

    /**
     * Handles real-time messages from the streaming service.
     *
     * This function listens for various types of messages from the `realtimeStreaming.messages()`
     * async iterator and processes them accordingly. The messages can indicate different events
     * such as session creation, audio transcript updates, audio data, speech start, transcription
     * completion, and the end of the response.
     *
     * The function performs the following actions based on the message type:
     * - "session.created": Updates the form input state, creates a new text block indicating the session start.
     * - "response.audio_transcript.delta": Appends the transcript delta to the current text block.
     * - "response.audio.delta": Decodes the audio data and plays it using the audio player.
     * - "input_audio_buffer.speech_started": Indicates the start of speech, prepares a new text block, and clears the audio player.
     * - "conversation.item.input_audio_transcription.completed": Appends the completed user transcription to the latest input speech block.
     * - "response.done": Appends a horizontal rule to the form received text container.
     * - Default: Logs the message as a JSON string.
     *
     * After processing all messages, the function resets the audio state.
     *
     * @async
     * @function handleRealtimeMessages
     * @returns {Promise<void>} A promise that resolves when the message handling is complete.
     */
    async function handleRealtimeMessages() {
        let question: string = "DEFAULT QUESTION";
        let answer: string = "";
        let askResponse: ChatAppResponse = {
            message: {} as ResponseMessage,
            delta: {} as ResponseMessage,
            context: { data_points: [], followup_questions: [], thoughts: [] },
            session_state: null
        } as ChatAppResponse;
        let tempMcqData = "";

        const updateAnswerState = (newContent: string) => {
            return new Promise(resolve => {
                setTimeout(() => {
                    answer += newContent;
                    const latestResponse: ChatAppResponse = {
                        ...askResponse,
                        message: { content: answer, role: askResponse.message.role }
                    };
                    console.log("updateAnswerState: setting streamed answers...");
                    console.log("updateAnswerState: answers: " + JSON.stringify(answers));
                    console.log("updateAnswerState: question: " + question);
                    setStreamedAnswers([...answers, [question, latestResponse]]);
                    resolve(null);
                }, 33);
            });
        };

        try {
            for await (const message of realtimeStreaming.messages()) {
                let consoleLog = "" + message.type;
                // console.log("handleRealtimeMessages: message: " + JSON.stringify(message));
                switch (message.type) {
                    case "session.created":
                        console.log("handleRealtimeMessages: 'Session created' sequence started...");
                        setFormInputState(InputState.ReadyToStop);
                        // makeNewTextBlock("<< Session Started >>");
                        // makeNewTextBlock();
                        askResponse = {
                            message: {} as ResponseMessage,
                            delta: {} as ResponseMessage,
                            context: { data_points: [], followup_questions: [], thoughts: [] },
                            session_state: null
                        } as ChatAppResponse;
                        console.log("handleRealtimeMessages: 'Session created' sequence ended.");
                        break;
                    case "response.audio_transcript.delta":
                        console.log("handleRealtimeMessages: Appending transcript delta: " + message.delta);
                        // if (message.delta.includes('{"')) {

                        tempMcqData += message.delta;
                        console.log("handleRealtimeMessages: temp prompt: " + tempMcqData);

                        //     console.log("handleRealtimeMessages: MCQ start detected: " + tempMcqData);
                        // } else if (message.delta.includes("}}")) {
                        //     tempMcqData += message.delta + "</mcqdata>";
                        //     setMcqPrompt(tempMcqData);
                        //     //displayPromptRendering();
                        //     console.log("handleRealtimeMessages: MCQ end detected: " + mcqPrompt);
                        // } else {
                        //     appendToTextBlock(message.delta);
                        //     updateAnswerState(message.delta);
                        //     console.log("handleRealtimeMessages: Transcript delta appended.");
                        // }
                        // appendToTextBlock(message.delta);
                        break;
                    case "response.audio.delta":
                        console.log("handleRealtimeMessages: Playing audio...");
                        const binary = atob(message.delta);
                        const bytes = Uint8Array.from(binary, c => c.charCodeAt(0));
                        const pcmData = new Int16Array(bytes.buffer);
                        audioPlayer.play(pcmData);
                        console.log("handleRealtimeMessages: Audio played.");
                        break;

                    case "input_audio_buffer.speech_started":
                        console.log("handleRealtimeMessages: 'Speech started' sequence started...");
                        //makeNewTextBlock("<< Speech Started >>");
                        let textElements = formReceivedTextContainer.children;
                        //latestInputSpeechBlock = textElements[textElements.length - 1];
                        //makeNewTextBlock();
                        audioPlayer.clear();
                        tempMcqData = "";

                        //displayPromptRendering();
                        console.log("handleRealtimeMessages: 'Speech started' sequence ended.");
                        break;
                    case "conversation.item.input_audio_transcription.completed":
                        console.log("handleRealtimeMessages: Appending completed user transcription...");
                        //latestInputSpeechBlock.textContent += "Nicolas: " + message.transcript;
                        question = message.transcript;
                        setStudentTurn("Nicolas: " + question);
                        setMcqPrompt("Greg: ...");

                        console.log("handleRealtimeMessages: Completed user transcription appended.");
                        break;
                    case "response.done":
                        console.log("handleRealtimeMessages: 'Response done' sequence started...");
                        //formReceivedTextContainer.appendChild(document.createElement("hr"));
                        console.log("handleRealtimeMessages: 'Response done' sequence ended.");
                        break;
                    case "response.audio_transcript.done":
                        console.log("handleRealtimeMessages: 'response.audio_transcript.done' sequence started...");
                        //displayPromptRendering();
                        //let textElements1 = formReceivedTextContainer.children;
                        //if (textElements1 && textElements1.length > 0) {
                        console.log("displayPromptRendering: switching to render prompt via PromptRenderer...");
                        // let lastElement = textElements1[textElements1.length - 1];
                        // let text = lastElement.textContent || "";
                        setMcqPrompt("Greg: " + tempMcqData);
                        //lastElement.textContent?.replace("<mcqdata>*</mcqdata>", "");
                        //}
                        console.log("handleRealtimeMessages: 'response.audio_transcript.done' sequence ended.");
                        break;
                    default:
                        console.log("handleRealtimeMessages: Default case. Logging message as JSON string...");
                        consoleLog = JSON.stringify(message, null, 2);
                        break;
                }
                if (consoleLog) {
                    console.log(consoleLog);
                }
            }
        } catch (error) {
            console.error("handleRealtimeMessages: Error occurred while handling messages:", error);
        } finally {
            console.log("handleRealtimeMessages: Realtime message handling complete. Resetting audio...");
            await resetAudio(false);
            setIsStreaming(false);
            console.log("handleRealtimeMessages: Audio reset.");
        }
        const fullResponse: ChatAppResponse = {
            ...askResponse,
            message: { content: answer, role: askResponse.message.role }
        };
        console.log("handleRealtimeMessages: Setting non-streamed answers...");
        console.log("handleRealtimeMessages: answers: " + JSON.stringify(answers));
        console.log("handleRealtimeMessages: question: " + question);
        setAnswers([...answers, [question, fullResponse]]);
    }

    /**
     * Basic audio handling
     */

    let recordingActive: boolean = false;
    let buffer: Uint8Array = new Uint8Array();

    /**
     * Combines the existing buffer with new data and updates the buffer.
     *
     * @param newData - The new data to be appended to the existing buffer.
     */
    function combineArray(newData: Uint8Array) {
        const newBuffer = new Uint8Array(buffer.length + newData.length);
        newBuffer.set(buffer);
        newBuffer.set(newData, buffer.length);
        buffer = newBuffer;
    }

    /**
     * Processes an audio recording buffer by converting it to a Uint8Array,
     * combining it with an existing buffer, and sending it in chunks if the buffer
     * length exceeds a specified threshold.
     *
     * @param data - The audio recording buffer to process.
     */
    function processAudioRecordingBuffer(data: Buffer) {
        const uint8Array = new Uint8Array(data);
        combineArray(uint8Array);
        if (buffer.length >= 4800) {
            const toSend = new Uint8Array(buffer.slice(0, 4800));
            buffer = new Uint8Array(buffer.slice(4800));
            const regularArray = String.fromCharCode(...toSend);
            const base64 = btoa(regularArray);
            if (recordingActive) {
                realtimeStreaming.send({
                    type: "input_audio_buffer.append",
                    audio: base64
                });
            }
        }
    }

    /**
     * Resets the audio recorder and player, and optionally starts a new recording session.
     *
     * @param {boolean} startRecording - If true, starts a new recording session after resetting.
     * @returns {Promise<void>} A promise that resolves when the audio has been reset and optionally started recording.
     *
     * @remarks
     * This function stops any active recording, clears the audio player, reinitializes the recorder and player,
     * and optionally starts a new recording session if `startRecording` is true.
     *
     * @example
     * ```typescript
     * // To reset audio and start a new recording session
     * await resetAudio(true);
     *
     * // To reset audio without starting a new recording session
     * await resetAudio(false);
     * ```
     */
    async function resetAudio(startRecording: boolean) {
        console.log("resetAudio: Acquiring lock");
        await lock.acquire();
        try {
            console.log("resetAudio: Lock acquired. Switching recorting active to false...");
            recordingActive = false;
            if (audioRecorder) {
                console.log("resetAudio: Stopping audio recorder...");
                audioRecorder.stop();
            }
            if (audioPlayer) {
                console.log("resetAudio: Clearing audio player...");
                audioPlayer.clear();
            }
            console.log("resetAudio: Reinitializing audio recorder and player...");
            audioRecorder = new Recorder(processAudioRecordingBuffer);
            audioPlayer = new Player();
            await audioPlayer.init(24000);
            if (startRecording) {
                console.log("resetAudio: Starting new recording session");
                const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
                audioRecorder.start(stream);
                recordingActive = true;
            }
        } catch (error) {
            console.error("resetAudio: Error during audio reset:", error);
        } finally {
            console.log("resetAudio: Audio reset process completed");
            lock.release();
            console.log("resetAudio: Lock released");
        }
    }

    /**
     * UI and controls
     */

    let formReceivedTextContainer = document.querySelector<HTMLDivElement>("#received-text-container")!;
    let formStartButton = document.querySelector<HTMLButtonElement>("#start-recording")!;
    let formStopButton = document.querySelector<HTMLButtonElement>("#stop-recording")!;
    //const formEndpointField = document.querySelector<HTMLInputElement>("#endpoint")!;
    //const formAzureToggle = document.querySelector<HTMLInputElement>("#azure-toggle")!;
    //const formApiKeyField = document.querySelector<HTMLInputElement>("#api-key")!;
    //const formDeploymentOrModelField = document.querySelector<HTMLInputElement>("#deployment-or-model")!;
    //const formSessionInstructionsField = document.querySelector<HTMLTextAreaElement>("#session-instructions")!;
    //const formTemperatureField = document.querySelector<HTMLInputElement>("#temperature")!;
    //const formVoiceSelection = document.querySelector<HTMLInputElement>("#voice")!;

    let latestInputSpeechBlock: Element;

    enum InputState {
        Working,
        ReadyToStart,
        ReadyToStop
    }

    function setFormInputState(state: InputState) {
        formStartButton.disabled = state != InputState.ReadyToStart;
        formStopButton.disabled = state != InputState.ReadyToStop;
    }

    function getSystemMessage(): string {
        console.log("getSystemMessage: returning prompt template: " + prmpt);
        if (prmpt && prmpt.length > 0) {
            return prmpt;
        }
        return SYSTEM_PROMPT.replace("{{{language}}}", i18n.language);
    }

    function getTemperature(): number {
        return TEMPERATURE;
    }

    function getVoice(): "alloy" | "echo" | "shimmer" {
        return VOICE as "alloy" | "echo" | "shimmer";
    }

    /**
     * Creates a new paragraph element with the specified text content and appends it to the formReceivedTextContainer.
     *
     * @param text - The text content to be added to the new paragraph element. Defaults to an empty string.
     */
    // function makeNewTextBlock(text: string = "") {
    //     let newElement = document.createElement("p");
    //     newElement.textContent = text;
    //     formReceivedTextContainer.appendChild(newElement);
    // }

    /**
     * Appends the given text to the last text block within the formReceivedTextContainer.
     * If there are no text blocks, a new one is created.
     *
     * @param text - The text to append to the last text block.
     */
    // function appendToTextBlock(text: string) {
    //     let textElements = formReceivedTextContainer.children;
    //     if (textElements.length == 0) {
    //         makeNewTextBlock();
    //     }
    //     textElements[textElements.length - 1].textContent += text;
    // }

    // function displayPromptRendering() {
    //     let textElements = formReceivedTextContainer.children;
    //     if (textElements.length > 0) {
    //         console.log("displayPromptRendering: switching to render prompt via PromptRenderer...");
    //         let lastElement = textElements[textElements.length - 1];
    //         let text = lastElement.textContent || "";
    //         setMcqPrompt(text);
    //         //lastElement.textContent?.replace("<mcqdata>*</mcqdata>", "");
    //     }
    // }

    useEffect(() => {
        formReceivedTextContainer = document.querySelector<HTMLDivElement>("#received-text-container")!;
        formStartButton = document.querySelector<HTMLButtonElement>("#start-recording")!;
        formStopButton = document.querySelector<HTMLButtonElement>("#stop-recording")!;

        //if (formStartButton) {
        formStartButton.addEventListener("click", async () => {
            setFormInputState(InputState.Working);

            try {
                makeApiRequest("INITIAL QUESTION");
            } catch (error) {
                console.log(error);
                setFormInputState(InputState.ReadyToStart);
            }

            // Cleanup event listeners on component unmount
            return () => {
                if (formStartButton) {
                    formStartButton.removeEventListener("click", () => {});
                }
                if (formStopButton) {
                    formStopButton.removeEventListener("click", () => {});
                }
            };
        });
        //}

        // if (formStopButton) {
        formStopButton.addEventListener("click", async () => {
            setFormInputState(InputState.Working);
            resetAudio(false);
            realtimeStreaming.close(); // !! IMPORTANT, Close the connection
            setFormInputState(InputState.ReadyToStart);
        });
        // }
    }, []);
    //
    //
    //
    //
    //
    // --------------------------------------------------------------------------
    // LEGACY CODE
    // --------------------------------------------------------------------------
    //
    //
    //
    //
    //
    const speechConfig: SpeechConfig = {
        speechUrls,
        setSpeechUrls,
        audio,
        isPlaying,
        setIsPlaying
    };

    const getConfig = async () => {
        configApi().then(config => {
            setShowGPT4VOptions(config.showGPT4VOptions);
            setUseSemanticRanker(config.showSemanticRankerOption);
            setShowSemanticRankerOption(config.showSemanticRankerOption);
            setShowVectorOption(config.showVectorOption);
            if (!config.showVectorOption) {
                setRetrievalMode(RetrievalMode.Text);
            }
            setShowUserUpload(false);
            setshowLanguagePicker(config.showLanguagePicker);
            setShowSpeechInput(config.showSpeechInput);
            setShowSpeechOutputBrowser(config.showSpeechOutputBrowser);
            setShowSpeechOutputAzure(config.showSpeechOutputAzure);
        });
    };

    const handleAsyncRequest = async (question: string, answers: [string, ChatAppResponse][], responseBody: ReadableStream<any>) => {
        let answer: string = "";
        let askResponse: ChatAppResponse = {} as ChatAppResponse;

        const updateState = (newContent: string) => {
            return new Promise(resolve => {
                setTimeout(() => {
                    answer += newContent;
                    const latestResponse: ChatAppResponse = {
                        ...askResponse,
                        message: { content: answer, role: askResponse.message.role }
                    };
                    setStreamedAnswers([...answers, [question, latestResponse]]);
                    resolve(null);
                }, 33);
            });
        };
        try {
            setIsStreaming(true);
            for await (const event of readNDJSONStream(responseBody)) {
                if (event["context"] && event["context"]["data_points"]) {
                    event["message"] = event["delta"];
                    askResponse = event as ChatAppResponse;
                } else if (event["delta"] && event["delta"]["content"]) {
                    setIsLoading(false);
                    await updateState(event["delta"]["content"]);
                } else if (event["context"]) {
                    // Update context with new keys from latest event
                    askResponse.context = { ...askResponse.context, ...event["context"] };
                } else if (event["error"]) {
                    throw Error(event["error"]);
                }
            }
        } finally {
            setIsStreaming(false);
        }
        const fullResponse: ChatAppResponse = {
            ...askResponse,
            message: { content: answer, role: askResponse.message.role }
        };
        return fullResponse;
    };

    const client = useLogin ? useMsal().instance : undefined;
    const { loggedIn } = useContext(LoginContext);

    const makeApiRequest = async (question: string) => {
        lastQuestionRef.current = question;

        error && setError(undefined);
        setIsLoading(true);
        setActiveCitation(undefined);
        setActiveAnalysisPanelTab(undefined);

        const token = client ? await getToken(client) : undefined;

        try {
            // const messages: ResponseMessage[] = answers.flatMap(a => [
            //     { content: a[0], role: "user" },
            //     { content: a[1].message.content, role: "assistant" }
            // ]);

            // const request: ChatAppRequest = {
            //     messages: [...messages, { content: question, role: "user" }],
            //     context: {
            //         overrides: {
            //             prompt_template: promptTemplate.length === 0 ? undefined : promptTemplate,
            //             exclude_category: excludeCategory.length === 0 ? undefined : excludeCategory,
            //             top: retrieveCount,
            //             temperature: temperature,
            //             minimum_reranker_score: minimumRerankerScore,
            //             minimum_search_score: minimumSearchScore,
            //             retrieval_mode: retrievalMode,
            //             semantic_ranker: useSemanticRanker,
            //             semantic_captions: useSemanticCaptions,
            //             suggest_followup_questions: useSuggestFollowupQuestions,
            //             use_oid_security_filter: useOidSecurityFilter,
            //             use_groups_security_filter: useGroupsSecurityFilter,
            //             vector_fields: vectorFieldList,
            //             use_gpt4v: useGPT4V,
            //             gpt4v_input: gpt4vInput,
            //             language: i18n.language,
            //             ...(seed !== null ? { seed: seed } : {})
            //         }
            //     },
            //     // AI Chat Protocol: Client must pass on any session state received from the server
            //     session_state: answers.length ? answers[answers.length - 1][1].session_state : null
            // };

            //const response = await chatApi(request, shouldStream, token);
            // if (!response.body) {
            //     throw Error("No response body");
            // }
            // if (response.status > 299 || !response.ok) {
            //     throw Error(`Request failed with status ${response.status}`);
            // }
            if (shouldStream) {
                start_realtime();
                //const parsedResponse: ChatAppResponse = await handleAsyncRequest(question, answers, response.body);
                //setAnswers([...answers, [question, parsedResponse]]);
                setWorkflowStateNo(workflowStateNo + 1);
            } else {
                // const parsedResponse: ChatAppResponseOrError = await response.json();
                // if (parsedResponse.error) {
                //     throw Error(parsedResponse.error);
                // }
                // setAnswers([...answers, [question, parsedResponse as ChatAppResponse]]);
                // setWorkflowStateNo(workflowStateNo + 1);
            }
            setSpeechUrls([...speechUrls, null]);
        } catch (e) {
            setError(e);
        } finally {
            setIsLoading(false);
        }
    };

    // const makeApiRequestLegacy = async (question: string) => {
    //     lastQuestionRef.current = question;

    //     error && setError(undefined);
    //     setIsLoading(true);
    //     setActiveCitation(undefined);
    //     setActiveAnalysisPanelTab(undefined);

    //     const token = client ? await getToken(client) : undefined;

    //     try {
    //         const messages: ResponseMessage[] = answers.flatMap(a => [
    //             { content: a[0], role: "user" },
    //             { content: a[1].message.content, role: "assistant" }
    //         ]);

    //         const request: ChatAppRequest = {
    //             messages: [...messages, { content: question, role: "user" }],
    //             context: {
    //                 overrides: {
    //                     prompt_template: promptTemplate.length === 0 ? undefined : promptTemplate,
    //                     exclude_category: excludeCategory.length === 0 ? undefined : excludeCategory,
    //                     top: retrieveCount,
    //                     temperature: temperature,
    //                     minimum_reranker_score: minimumRerankerScore,
    //                     minimum_search_score: minimumSearchScore,
    //                     retrieval_mode: retrievalMode,
    //                     semantic_ranker: useSemanticRanker,
    //                     semantic_captions: useSemanticCaptions,
    //                     suggest_followup_questions: useSuggestFollowupQuestions,
    //                     use_oid_security_filter: useOidSecurityFilter,
    //                     use_groups_security_filter: useGroupsSecurityFilter,
    //                     vector_fields: vectorFieldList,
    //                     use_gpt4v: useGPT4V,
    //                     gpt4v_input: gpt4vInput,
    //                     language: i18n.language,
    //                     ...(seed !== null ? { seed: seed } : {})
    //                 }
    //             },
    //             // AI Chat Protocol: Client must pass on any session state received from the server
    //             session_state: answers.length ? answers[answers.length - 1][1].session_state : null
    //         };

    //         const response = await chatApi(request, shouldStream, token);
    //         if (!response.body) {
    //             throw Error("No response body");
    //         }
    //         if (response.status > 299 || !response.ok) {
    //             throw Error(`Request failed with status ${response.status}`);
    //         }
    //         if (shouldStream) {
    //             const parsedResponse: ChatAppResponse = await handleAsyncRequest(question, answers, response.body);
    //             setAnswers([...answers, [question, parsedResponse]]);
    //             setWorkflowStateNo(workflowStateNo + 1);
    //         } else {
    //             const parsedResponse: ChatAppResponseOrError = await response.json();
    //             if (parsedResponse.error) {
    //                 throw Error(parsedResponse.error);
    //             }
    //             setAnswers([...answers, [question, parsedResponse as ChatAppResponse]]);
    //             setWorkflowStateNo(workflowStateNo + 1);
    //         }
    //         setSpeechUrls([...speechUrls, null]);
    //     } catch (e) {
    //         setError(e);
    //     } finally {
    //         setIsLoading(false);
    //     }
    // };

    const clearChat = () => {
        lastQuestionRef.current = "";
        error && setError(undefined);
        setActiveCitation(undefined);
        setActiveAnalysisPanelTab(undefined);
        setAnswers([]);
        setSpeechUrls([]);
        setStreamedAnswers([]);
        setIsLoading(false);
        setIsStreaming(false);
    };

    console.log("Workflow State No: ", workflowStateNo);

    useEffect(() => chatMessageStreamEnd.current?.scrollIntoView({ behavior: "smooth" }), [isLoading]);
    useEffect(() => chatMessageStreamEnd.current?.scrollIntoView({ behavior: "auto" }), [streamedAnswers]);
    useEffect(() => {
        getConfig();
    }, []);

    const onPromptTemplateChange = (_ev?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        console.log("onPromptTemplateChange: newValue: " + newValue);
        setPromptTemplate(newValue || "");
        prmpt = newValue || "";
        console.log("onPromptTemplateChange: promptTemplate: " + promptTemplate);
    };

    const onTemperatureChange = (_ev?: React.SyntheticEvent<HTMLElement, Event>, newValue?: string) => {
        setTemperature(parseFloat(newValue || "0"));
    };

    const onSeedChange = (_ev?: React.SyntheticEvent<HTMLElement, Event>, newValue?: string) => {
        setSeed(parseInt(newValue || ""));
    };

    const onMinimumSearchScoreChange = (_ev?: React.SyntheticEvent<HTMLElement, Event>, newValue?: string) => {
        setMinimumSearchScore(parseFloat(newValue || "0"));
    };

    const onMinimumRerankerScoreChange = (_ev?: React.SyntheticEvent<HTMLElement, Event>, newValue?: string) => {
        setMinimumRerankerScore(parseFloat(newValue || "0"));
    };

    const onRetrieveCountChange = (_ev?: React.SyntheticEvent<HTMLElement, Event>, newValue?: string) => {
        setRetrieveCount(parseInt(newValue || "3"));
    };

    const onUseSemanticRankerChange = (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
        setUseSemanticRanker(!!checked);
    };

    const onUseSemanticCaptionsChange = (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
        setUseSemanticCaptions(!!checked);
    };

    const onShouldStreamChange = (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
        setShouldStream(!!checked);
    };

    const onExcludeCategoryChanged = (_ev?: React.FormEvent, newValue?: string) => {
        setExcludeCategory(newValue || "");
    };

    const onUseSuggestFollowupQuestionsChange = (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
        setUseSuggestFollowupQuestions(!!checked);
    };

    const onUseOidSecurityFilterChange = (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
        setUseOidSecurityFilter(!!checked);
    };

    const onUseGroupsSecurityFilterChange = (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
        setUseGroupsSecurityFilter(!!checked);
    };

    const onExampleClicked = (example: string) => {
        makeApiRequest(example);
    };

    const onShowCitation = (citation: string, index: number) => {
        if (activeCitation === citation && activeAnalysisPanelTab === AnalysisPanelTabs.CitationTab && selectedAnswer === index) {
            setActiveAnalysisPanelTab(undefined);
        } else {
            setActiveCitation(citation);
            setActiveAnalysisPanelTab(AnalysisPanelTabs.CitationTab);
        }

        setSelectedAnswer(index);
    };

    const onToggleTab = (tab: AnalysisPanelTabs, index: number) => {
        if (activeAnalysisPanelTab === tab && selectedAnswer === index) {
            setActiveAnalysisPanelTab(undefined);
        } else {
            setActiveAnalysisPanelTab(tab);
        }

        setSelectedAnswer(index);
    };

    // IDs for form labels and their associated callouts
    const promptTemplateId = useId("promptTemplate");
    const promptTemplateFieldId = useId("promptTemplateField");
    const temperatureId = useId("temperature");
    const temperatureFieldId = useId("temperatureField");
    const seedId = useId("seed");
    const seedFieldId = useId("seedField");
    const searchScoreId = useId("searchScore");
    const searchScoreFieldId = useId("searchScoreField");
    const rerankerScoreId = useId("rerankerScore");
    const rerankerScoreFieldId = useId("rerankerScoreField");
    const retrieveCountId = useId("retrieveCount");
    const retrieveCountFieldId = useId("retrieveCountField");
    const excludeCategoryId = useId("excludeCategory");
    const excludeCategoryFieldId = useId("excludeCategoryField");
    const semanticRankerId = useId("semanticRanker");
    const semanticRankerFieldId = useId("semanticRankerField");
    const semanticCaptionsId = useId("semanticCaptions");
    const semanticCaptionsFieldId = useId("semanticCaptionsField");
    const suggestFollowupQuestionsId = useId("suggestFollowupQuestions");
    const suggestFollowupQuestionsFieldId = useId("suggestFollowupQuestionsField");
    const useOidSecurityFilterId = useId("useOidSecurityFilter");
    const useOidSecurityFilterFieldId = useId("useOidSecurityFilterField");
    const useGroupsSecurityFilterId = useId("useGroupsSecurityFilter");
    const useGroupsSecurityFilterFieldId = useId("useGroupsSecurityFilterField");
    const shouldStreamId = useId("shouldStream");
    const shouldStreamFieldId = useId("shouldStreamField");
    const { t, i18n } = useTranslation();

    return (
        <div className={styles.container}>
            {/* Setting the page title using react-helmet-async */}
            <Helmet>
                <title>{t("pageTitle")}</title>
            </Helmet>

            <div className={styles.commandsContainer}>
                {/*<ClearChatButton className={styles.commandButton} onClick={clearChat} disabled={!lastQuestionRef.current || isLoading} />*/}
                {showUserUpload && <UploadFile className={styles.commandButton} disabled={!loggedIn} />}
                <SettingsButton className={styles.commandButton} onClick={() => setIsConfigPanelOpen(!isConfigPanelOpen)} />
            </div>
            <div className={styles.chatRoot}>
                <div className={styles.chatContainer}>
                    <div id="received-text-container" className={styles.chatMessageGptMinWidth}></div>
                    <PromptRenderer prompt={studentTurn} />
                    <PromptRenderer prompt={mcqPrompt} />
                    {!lastQuestionRef.current ? (
                        <div className={styles.chatEmptyState}>
                            <SparkleFilled fontSize={"120px"} primaryFill={"rgba(115, 118, 225, 1)"} aria-hidden="true" aria-label="Chat logo" />
                            <h1 className={styles.chatEmptyStateTitle}>{t("chatEmptyStateTitle")}</h1>
                            <h2 className={styles.chatEmptyStateSubtitle}>Hello Nicolas! Welcome to our Tutoring session!</h2>
                            <hr />
                            {showLanguagePicker && <LanguagePicker onLanguageChange={newLang => i18n.changeLanguage(newLang)} />}

                            <ExampleList onExampleClicked={onExampleClicked} useGPT4V={useGPT4V} />
                        </div>
                    ) : (
                        <div className={styles.chatMessageStream}>
                            {/* {isStreaming && workflowStateNo > 0 && (
                                <div key={0}>
                                    <UserChatMessage message={streamedAnswers[0] ? streamedAnswers[0][0] : ""} />
                                    <div className={styles.chatMessageGpt}>
                                        <Answer
                                            isStreaming={true}
                                            key="0"
                                            answer={
                                                streamedAnswers[streamedAnswers.length - 1] &&
                                                typeof streamedAnswers[streamedAnswers.length - 1][1] !== "string"
                                                    ? streamedAnswers[streamedAnswers.length - 1][1]
                                                    : ({} as ChatAppResponse)
                                            }
                                            index={streamedAnswers.length - 1}
                                            speechConfig={speechConfig}
                                            isSelected={false}
                                            onCitationClicked={c => onShowCitation(c, 0)}
                                            // onThoughtProcessClicked={() => onToggleTab(AnalysisPanelTabs.ThoughtProcessTab, index)}
                                            onSupportingContentClicked={() => onToggleTab(AnalysisPanelTabs.SupportingContentTab, 0)}
                                            onFollowupQuestionClicked={q => makeApiRequest(q)}
                                            showFollowupQuestions={useSuggestFollowupQuestions && answers.length - 1 === 0}
                                            showSpeechOutputAzure={showSpeechOutputAzure}
                                            showSpeechOutputBrowser={showSpeechOutputBrowser}
                                            workflowStateNo={workflowStateNo}
                                            onMCQAnswerSelected={handleMCQAnswerSelected}
                                        />
                                    </div>
                                </div>
                            )} */}
                            {/* {isStreaming &&
                                streamedAnswers.map((streamedAnswer, index) => (
                                    <div key={index}>
                                        <UserChatMessage message={streamedAnswer[0]} />
                                        <div className={styles.chatMessageGpt}>
                                            <Answer
                                                isStreaming={true}
                                                key={index}
                                                answer={streamedAnswer[1]}
                                                index={index}
                                                speechConfig={speechConfig}
                                                isSelected={false}
                                                onCitationClicked={c => onShowCitation(c, index)}
                                                // onThoughtProcessClicked={() => onToggleTab(AnalysisPanelTabs.ThoughtProcessTab, index)}
                                                onSupportingContentClicked={() => onToggleTab(AnalysisPanelTabs.SupportingContentTab, index)}
                                                onFollowupQuestionClicked={q => makeApiRequest(q)}
                                                showFollowupQuestions={useSuggestFollowupQuestions && answers.length - 1 === index}
                                                showSpeechOutputAzure={showSpeechOutputAzure}
                                                showSpeechOutputBrowser={showSpeechOutputBrowser}
                                                workflowStateNo={workflowStateNo}
                                            />
                                        </div>
                                    </div>
                                ))} */}
                            {/* {!isStreaming && workflowStateNo > 0 && (
                                <div key={0}>
                                    <UserChatMessage message={answers[0] ? answers[0][0] : ""} />
                                    <div className={styles.chatMessageGpt}>
                                        <Answer
                                            isStreaming={false}
                                            key={answers.length - 1}
                                            answer={
                                                answers[answers.length - 1] && typeof answers[answers.length - 1][1] !== "string"
                                                    ? answers[answers.length - 1][1]
                                                    : ({} as ChatAppResponse)
                                            }
                                            index={answers.length - 1}
                                            speechConfig={speechConfig}
                                            isSelected={selectedAnswer === 0 && activeAnalysisPanelTab !== undefined}
                                            onCitationClicked={c => onShowCitation(c, 0)}
                                            // onThoughtProcessClicked={() => onToggleTab(AnalysisPanelTabs.ThoughtProcessTab, index)}
                                            onSupportingContentClicked={() => onToggleTab(AnalysisPanelTabs.SupportingContentTab, 0)}
                                            onFollowupQuestionClicked={q => makeApiRequest(q)}
                                            showFollowupQuestions={useSuggestFollowupQuestions && answers.length - 1 === 0}
                                            showSpeechOutputAzure={showSpeechOutputAzure}
                                            showSpeechOutputBrowser={showSpeechOutputBrowser}
                                            workflowStateNo={workflowStateNo}
                                            onMCQAnswerSelected={handleMCQAnswerSelected}
                                        />
                                    </div>
                                </div>
                            )} */}
                            {/* {!isStreaming &&
                                answers.map((answer, index) => (
                                    <div key={index}>
                                        <UserChatMessage message={answer[0]} />
                                        <div className={styles.chatMessageGpt}>
                                            <Answer
                                                isStreaming={false}
                                                key={index}
                                                answer={answer[1]}
                                                index={index}
                                                speechConfig={speechConfig}
                                                isSelected={selectedAnswer === index && activeAnalysisPanelTab !== undefined}
                                                onCitationClicked={c => onShowCitation(c, index)}
                                                // onThoughtProcessClicked={() => onToggleTab(AnalysisPanelTabs.ThoughtProcessTab, index)}
                                                onSupportingContentClicked={() => onToggleTab(AnalysisPanelTabs.SupportingContentTab, index)}
                                                onFollowupQuestionClicked={q => makeApiRequest(q)}
                                                showFollowupQuestions={useSuggestFollowupQuestions && answers.length - 1 === index}
                                                showSpeechOutputAzure={showSpeechOutputAzure}
                                                showSpeechOutputBrowser={showSpeechOutputBrowser}
                                                workflowStateNo={workflowStateNo}
                                            />
                                        </div>
                                    </div>
                                ))} */}
                            {isLoading && (
                                <>
                                    <UserChatMessage message={lastQuestionRef.current} />
                                    <div className={styles.chatMessageGptMinWidth}>
                                        <AnswerLoading />
                                    </div>
                                </>
                            )}
                            {error ? (
                                <>
                                    <UserChatMessage message={lastQuestionRef.current} />
                                    <div className={styles.chatMessageGptMinWidth}>
                                        <AnswerError error={error.toString()} onRetry={() => makeApiRequest(lastQuestionRef.current)} />
                                    </div>
                                </>
                            ) : null}
                            <div ref={chatMessageStreamEnd} />
                        </div>
                    )}

                    {workflowStateNo < 9 && (
                        <div className={styles.chatInput}>
                            <QuestionInput
                                clearOnSend
                                placeholder={t("defaultExamples.placeholder")}
                                disabled={isLoading}
                                onSend={question => makeApiRequest(question)}
                                showSpeechInput={showSpeechInput}
                                initQuestion={selectedMCQAnswer ? "My answer is: " + selectedMCQAnswer : ""}
                            />
                        </div>
                    )}
                </div>

                {answers.length > 0 && activeAnalysisPanelTab && (
                    <AnalysisPanel
                        className={styles.chatAnalysisPanel}
                        activeCitation={activeCitation}
                        onActiveTabChanged={x => onToggleTab(x, selectedAnswer)}
                        citationHeight="810px"
                        answer={answers[selectedAnswer][1]}
                        activeTab={activeAnalysisPanelTab}
                    />
                )}

                <Panel
                    headerText={t("labels.headerText")}
                    isOpen={isConfigPanelOpen}
                    isBlocking={false}
                    onDismiss={() => setIsConfigPanelOpen(false)}
                    closeButtonAriaLabel={t("labels.closeButton")}
                    onRenderFooterContent={() => <DefaultButton onClick={() => setIsConfigPanelOpen(false)}>{t("labels.closeButton")}</DefaultButton>}
                    isFooterAtBottom={true}
                >
                    <TextField
                        id={promptTemplateFieldId}
                        className={styles.chatSettingsSeparator}
                        defaultValue={promptTemplate}
                        label={t("labels.promptTemplate")}
                        multiline
                        autoAdjustHeight
                        onChange={onPromptTemplateChange}
                        aria-labelledby={promptTemplateId}
                        onRenderLabel={(props: ITextFieldProps | undefined) => (
                            <HelpCallout
                                labelId={promptTemplateId}
                                fieldId={promptTemplateFieldId}
                                helpText={t("helpTexts.promptTemplate")}
                                label={props?.label}
                            />
                        )}
                    />

                    <TextField
                        id={temperatureFieldId}
                        className={styles.chatSettingsSeparator}
                        label={t("labels.temperature")}
                        type="number"
                        min={0}
                        max={1}
                        step={0.1}
                        defaultValue={temperature.toString()}
                        onChange={onTemperatureChange}
                        aria-labelledby={temperatureId}
                        onRenderLabel={(props: ITextFieldProps | undefined) => (
                            <HelpCallout labelId={temperatureId} fieldId={temperatureFieldId} helpText={t("helpTexts.temperature")} label={props?.label} />
                        )}
                    />

                    <TextField
                        id={seedFieldId}
                        className={styles.chatSettingsSeparator}
                        label={t("labels.seed")}
                        type="text"
                        defaultValue={seed?.toString() || ""}
                        onChange={onSeedChange}
                        aria-labelledby={seedId}
                        onRenderLabel={(props: ITextFieldProps | undefined) => (
                            <HelpCallout labelId={seedId} fieldId={seedFieldId} helpText={t("helpTexts.seed")} label={props?.label} />
                        )}
                    />

                    <TextField
                        id={searchScoreFieldId}
                        className={styles.chatSettingsSeparator}
                        label={t("labels.minimumSearchScore")}
                        type="number"
                        min={0}
                        step={0.01}
                        defaultValue={minimumSearchScore.toString()}
                        onChange={onMinimumSearchScoreChange}
                        aria-labelledby={searchScoreId}
                        onRenderLabel={(props: ITextFieldProps | undefined) => (
                            <HelpCallout labelId={searchScoreId} fieldId={searchScoreFieldId} helpText={t("helpTexts.searchScore")} label={props?.label} />
                        )}
                    />

                    {showSemanticRankerOption && (
                        <TextField
                            id={rerankerScoreFieldId}
                            className={styles.chatSettingsSeparator}
                            label={t("labels.minimumRerankerScore")}
                            type="number"
                            min={1}
                            max={4}
                            step={0.1}
                            defaultValue={minimumRerankerScore.toString()}
                            onChange={onMinimumRerankerScoreChange}
                            aria-labelledby={rerankerScoreId}
                            onRenderLabel={(props: ITextFieldProps | undefined) => (
                                <HelpCallout
                                    labelId={rerankerScoreId}
                                    fieldId={rerankerScoreFieldId}
                                    helpText={t("helpTexts.rerankerScore")}
                                    label={props?.label}
                                />
                            )}
                        />
                    )}

                    <TextField
                        id={retrieveCountFieldId}
                        className={styles.chatSettingsSeparator}
                        label={t("labels.retrieveCount")}
                        type="number"
                        min={1}
                        max={50}
                        defaultValue={retrieveCount.toString()}
                        onChange={onRetrieveCountChange}
                        aria-labelledby={retrieveCountId}
                        onRenderLabel={(props: ITextFieldProps | undefined) => (
                            <HelpCallout
                                labelId={retrieveCountId}
                                fieldId={retrieveCountFieldId}
                                helpText={t("helpTexts.retrieveNumber")}
                                label={props?.label}
                            />
                        )}
                    />

                    <TextField
                        id={excludeCategoryFieldId}
                        className={styles.chatSettingsSeparator}
                        label={t("labels.excludeCategory")}
                        defaultValue={excludeCategory}
                        onChange={onExcludeCategoryChanged}
                        aria-labelledby={excludeCategoryId}
                        onRenderLabel={(props: ITextFieldProps | undefined) => (
                            <HelpCallout
                                labelId={excludeCategoryId}
                                fieldId={excludeCategoryFieldId}
                                helpText={t("helpTexts.excludeCategory")}
                                label={props?.label}
                            />
                        )}
                    />

                    {showSemanticRankerOption && (
                        <>
                            <Checkbox
                                id={semanticRankerFieldId}
                                className={styles.chatSettingsSeparator}
                                checked={useSemanticRanker}
                                label={t("labels.useSemanticRanker")}
                                onChange={onUseSemanticRankerChange}
                                aria-labelledby={semanticRankerId}
                                onRenderLabel={(props: ICheckboxProps | undefined) => (
                                    <HelpCallout
                                        labelId={semanticRankerId}
                                        fieldId={semanticRankerFieldId}
                                        helpText={t("helpTexts.useSemanticReranker")}
                                        label={props?.label}
                                    />
                                )}
                            />

                            <Checkbox
                                id={semanticCaptionsFieldId}
                                className={styles.chatSettingsSeparator}
                                checked={useSemanticCaptions}
                                label={t("labels.useSemanticCaptions")}
                                onChange={onUseSemanticCaptionsChange}
                                disabled={!useSemanticRanker}
                                aria-labelledby={semanticCaptionsId}
                                onRenderLabel={(props: ICheckboxProps | undefined) => (
                                    <HelpCallout
                                        labelId={semanticCaptionsId}
                                        fieldId={semanticCaptionsFieldId}
                                        helpText={t("helpTexts.useSemanticCaptions")}
                                        label={props?.label}
                                    />
                                )}
                            />
                        </>
                    )}

                    <Checkbox
                        id={suggestFollowupQuestionsFieldId}
                        className={styles.chatSettingsSeparator}
                        checked={useSuggestFollowupQuestions}
                        label={t("labels.useSuggestFollowupQuestions")}
                        onChange={onUseSuggestFollowupQuestionsChange}
                        aria-labelledby={suggestFollowupQuestionsId}
                        onRenderLabel={(props: ICheckboxProps | undefined) => (
                            <HelpCallout
                                labelId={suggestFollowupQuestionsId}
                                fieldId={suggestFollowupQuestionsFieldId}
                                helpText={t("helpTexts.suggestFollowupQuestions")}
                                label={props?.label}
                            />
                        )}
                    />

                    {showGPT4VOptions && (
                        <GPT4VSettings
                            gpt4vInputs={gpt4vInput}
                            isUseGPT4V={useGPT4V}
                            updateUseGPT4V={useGPT4V => {
                                setUseGPT4V(useGPT4V);
                            }}
                            updateGPT4VInputs={inputs => setGPT4VInput(inputs)}
                        />
                    )}

                    {showVectorOption && (
                        <VectorSettings
                            defaultRetrievalMode={retrievalMode}
                            showImageOptions={useGPT4V && showGPT4VOptions}
                            updateVectorFields={(options: VectorFieldOptions[]) => setVectorFieldList(options)}
                            updateRetrievalMode={(retrievalMode: RetrievalMode) => setRetrievalMode(retrievalMode)}
                        />
                    )}

                    {useLogin && (
                        <>
                            <Checkbox
                                id={useOidSecurityFilterFieldId}
                                className={styles.chatSettingsSeparator}
                                checked={useOidSecurityFilter || requireAccessControl}
                                label={t("labels.useOidSecurityFilter")}
                                disabled={!loggedIn || requireAccessControl}
                                onChange={onUseOidSecurityFilterChange}
                                aria-labelledby={useOidSecurityFilterId}
                                onRenderLabel={(props: ICheckboxProps | undefined) => (
                                    <HelpCallout
                                        labelId={useOidSecurityFilterId}
                                        fieldId={useOidSecurityFilterFieldId}
                                        helpText={t("helpTexts.useOidSecurityFilter")}
                                        label={props?.label}
                                    />
                                )}
                            />
                            <Checkbox
                                id={useGroupsSecurityFilterFieldId}
                                className={styles.chatSettingsSeparator}
                                checked={useGroupsSecurityFilter || requireAccessControl}
                                label={t("labels.useGroupsSecurityFilter")}
                                disabled={!loggedIn || requireAccessControl}
                                onChange={onUseGroupsSecurityFilterChange}
                                aria-labelledby={useGroupsSecurityFilterId}
                                onRenderLabel={(props: ICheckboxProps | undefined) => (
                                    <HelpCallout
                                        labelId={useGroupsSecurityFilterId}
                                        fieldId={useGroupsSecurityFilterFieldId}
                                        helpText={t("helpTexts.useGroupsSecurityFilter")}
                                        label={props?.label}
                                    />
                                )}
                            />
                        </>
                    )}

                    <Checkbox
                        id={shouldStreamFieldId}
                        className={styles.chatSettingsSeparator}
                        checked={shouldStream}
                        label={t("labels.shouldStream")}
                        onChange={onShouldStreamChange}
                        aria-labelledby={shouldStreamId}
                        onRenderLabel={(props: ICheckboxProps | undefined) => (
                            <HelpCallout labelId={shouldStreamId} fieldId={shouldStreamFieldId} helpText={t("helpTexts.streamChat")} label={props?.label} />
                        )}
                    />

                    {useLogin && <TokenClaimsDisplay />}
                </Panel>
            </div>
        </div>
    );
};

export default Chat;
