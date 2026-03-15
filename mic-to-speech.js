import os from "os";
import { execFile } from "child_process";
import * as sdk from "microsoft-cognitiveservices-speech-sdk";
import 'dotenv/config';
import process from 'process';

const scripted = true;
const recordingDurationSec = 10;

const referenceText = 'Both, simultaneously and in very different contexts, as the net effect depends almost entirely on how deliberately and consciously they are being used by each individual person.';

// Record from the default microphone and return a WAV buffer (16 kHz, 16-bit, mono)
function recordMicToWavBuffer(durationSec) {
    return new Promise((resolve, reject) => {
        const platform = os.platform();

        let inputFormat, inputDevice;
        if (platform === 'darwin') {
            inputFormat = 'avfoundation';
            inputDevice = ':0';          // default microphone on macOS
        } else if (platform === 'win32') {
            inputFormat = 'dshow';
            inputDevice = 'audio=Microphone';
        } else {
            inputFormat = 'alsa';
            inputDevice = 'default';
        }

        const args = [
            '-f', inputFormat,
            '-i', inputDevice,
            '-t', String(durationSec),
            '-ar', '16000',
            '-ac', '1',
            '-f', 'wav',
            'pipe:1'
        ];

        execFile('ffmpeg', args, { encoding: 'buffer', maxBuffer: 50 * 1024 * 1024 }, (err, stdout, stderr) => {
            if (err) return reject(new Error(`ffmpeg error: ${stderr.toString()}`));
            resolve(stdout);
        });
    });
}

async function main() {
    const speechConfig = sdk.SpeechConfig.fromSubscription(
        process.env.SPEECH_KEY,
        process.env.SPEECH_REGION
    );
    speechConfig.speechRecognitionLanguage = process.env.LANGUAGE;

    console.log(`Recording from microphone for ${recordingDurationSec} seconds... (speak now)`);
    const wavBuffer = await recordMicToWavBuffer(recordingDurationSec);
    console.log(`Recorded: ${(wavBuffer.length / 1024).toFixed(1)} KB`);

    const audioConfig = sdk.AudioConfig.fromWavFileInput(wavBuffer);

    const pronunciationAssessmentConfig = new sdk.PronunciationAssessmentConfig(
        scripted ? referenceText : "",
        sdk.PronunciationAssessmentGradingSystem.HundredMark,
        sdk.PronunciationAssessmentGranularity.Phoneme,
        scripted ? true : false
    );
    pronunciationAssessmentConfig.enableProsodyAssessment = true;

    const reco = new sdk.SpeechRecognizer(speechConfig, audioConfig);

    reco.sessionStarted = (_s, e) => {
        console.log(`SESSION ID: ${e.sessionId}`);
    };
    pronunciationAssessmentConfig.applyTo(reco);

    reco.recognizeOnceAsync(
        async (result) => {
            try {
                const pronunciationResult = sdk.PronunciationAssessmentResult.fromResult(result);

                console.log(`Pronunciation assessment for: "${result.text}"`);
                console.log(` Accuracy score:      ${pronunciationResult.accuracyScore}`);
                console.log(` Pronunciation score: ${pronunciationResult.pronunciationScore}`);
                console.log(` Completeness score:  ${pronunciationResult.completenessScore}`);
                console.log(` Fluency score:       ${pronunciationResult.fluencyScore}`);
                console.log(` Prosody score:       ${pronunciationResult.prosodyScore}`);

                console.log("  Word-level details:");
                pronunciationResult.detailResult.Words.forEach((word, idx) => {
                    console.log(`    ${idx + 1}: word: ${word.Word}\taccuracy score: ${word.PronunciationAssessment.AccuracyScore}\terror type: ${word.PronunciationAssessment.ErrorType}`);
                });

                const output = {
                    referenceText,
                    recognizedText: result.text,
                    scores: {
                        accuracy: pronunciationResult.accuracyScore,
                        pronunciation: pronunciationResult.pronunciationScore,
                        completeness: pronunciationResult.completenessScore,
                        fluency: pronunciationResult.fluencyScore,
                        prosody: pronunciationResult.prosodyScore,
                    },
                    words: pronunciationResult.detailResult.Words.map((word) => ({
                        word: word.Word,
                        accuracyScore: word.PronunciationAssessment.AccuracyScore,
                        errorType: word.PronunciationAssessment.ErrorType,
                    })),
                };

                console.log(`\nResult: ${JSON.stringify(output, null, 2)}`);
            } catch (err) {
                console.error("Error processing result:", err);
            } finally {
                reco.close();
            }
        },
        (err) => {
            console.error("Recognition error:", err);
            reco.close();
        }
    );
}

main().catch((err) => {
    console.error("Fatal error:", err);
    process.exit(1);
});