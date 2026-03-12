import fs from "fs";
import path from "path";
import { execFile } from "child_process";
import * as XLSX from 'xlsx';
import * as sdk from "microsoft-cognitiveservices-speech-sdk";
import 'dotenv/config';
import process from 'process';

// Convert an MP3 file to a WAV buffer (16 kHz, 16-bit, mono) using ffmpeg
function mp3ToWavBuffer(filePath) {
    return new Promise((resolve, reject) => {
        execFile(
            'ffmpeg',
            ['-i', filePath, '-ar', '16000', '-ac', '1', '-f', 'wav', 'pipe:1'],
            { encoding: 'buffer', maxBuffer: 50 * 1024 * 1024 },
            (err, stdout, stderr) => {
                if (err) return reject(new Error(`ffmpeg error: ${stderr.toString()}`));
                resolve(stdout);
            }
        );
    });
}

const scripted = false;
const answerNumber = 1;
const fileName = '1.mp3';

const xlsxPath = 'xlsx/828_Answer.xlsx';
const recordedDir = 'recorded';
const resultDir = 'result';
const recordedFile = path.join(recordedDir, fileName);

async function main() {
    // Ensure output folders exist
    await fs.promises.mkdir(recordedDir, { recursive: true });
    await fs.promises.mkdir(resultDir, { recursive: true });

    // Load reference text from XLSX
    const workbook = XLSX.read(fs.readFileSync(xlsxPath));
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);
    const referenceText = jsonData.find(row => row.id === answerNumber)?.name;

    if (!referenceText) {
        console.error(`No text found for id ${answerNumber} in ${xlsxPath}`);
        process.exit(1);
    }

    // Check recorded file exists
    try {
        await fs.promises.access(recordedFile, fs.constants.F_OK);
    } catch {
        console.error(`Recorded file not found: ${recordedFile}`);
        process.exit(1);
    }

    const speechConfig = sdk.SpeechConfig.fromSubscription(
        process.env.SPEECH_KEY,
        process.env.SPEECH_REGION
    );
    speechConfig.speechRecognitionLanguage = process.env.LANGUAGE;

    console.log('Converting MP3 to WAV...');
    const wavBuffer = await mp3ToWavBuffer(recordedFile);
    console.log(`Converted: ${(wavBuffer.length / 1024).toFixed(1)} KB`);

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

                // Build result object
                const output = {
                    id: answerNumber,
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

                // Save result to result/<answerNumber>.json
                const resultFile = path.join(resultDir, `${answerNumber}-${fileName}.json`);
                await fs.promises.writeFile(resultFile, JSON.stringify(output, null, 2), 'utf-8');
                console.log(`\nResult saved to ${resultFile}`);
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