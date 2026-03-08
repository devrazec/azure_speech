import fs from "fs";
import path from "path";
import * as XLSX from 'xlsx';
import * as sdk from "microsoft-cognitiveservices-speech-sdk";
import 'dotenv/config';
import process from 'process';

//const xlsxPath = 'xlsx/276_Question.xlsx';
//const outputDir = 'mp3/question';

const xlsxPath = 'xlsx/828_Answer.xlsx';
const outputDir = 'mp3/answer';

// Ensure output folder exists
await fs.promises.mkdir(outputDir, { recursive: true });

// Helper: check if file exists
async function fileExists(filePath) {
    try {
        await fs.promises.access(filePath, fs.constants.F_OK);
        return true;
    } catch {
        return false;
    }
}

// Helper: synthesize text to an MP3 file
function synthesizeToFile(text, filePath) {
    return new Promise((resolve, reject) => {
        const speechConfig = sdk.SpeechConfig.fromSubscription(
            process.env.SPEECH_KEY,
            process.env.SPEECH_REGION
        );
        speechConfig.speechSynthesisVoiceName = "en-US-Emma2:DragonHDLatestNeural"; 
         
        // en-US-JennyNeural
        // en-US-RyanMultilingualNeural
        // en-US-BrianMultilingualNeural
        // en-US-Brian:DragonHDLatestNeural
        // en-US-Ava:DragonHDLatestNeural
        //
        // Dragon HD voices
        // en-US-Andrew2:DragonHDLatestNeural
        // en-US-Emma2:DragonHDLatestNeural

        speechConfig.speechSynthesisOutputFormat =
            sdk.SpeechSynthesisOutputFormat.Audio44Khz128KBitRateMonoMp3;

        const audioConfig = sdk.AudioConfig.fromAudioFileOutput(filePath);
        const synthesizer = new sdk.SpeechSynthesizer(speechConfig, audioConfig);

        synthesizer.speakTextAsync(
            text,
            result => {
                synthesizer.close();
                if (result.reason === sdk.ResultReason.SynthesizingAudioCompleted) {
                    resolve();
                } else {
                    reject(new Error(result.errorDetails ?? `Synthesis failed: ${result.reason}`));
                }
            },
            err => {
                synthesizer.close();
                reject(typeof err === 'string' ? new Error(err) : err);
            }
        );
    });
}

// Generate audio for all words sequentially
async function generateAudio() {

    const workbook = XLSX.read(fs.readFileSync(xlsxPath));
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    console.log(`📄 Loaded ${jsonData.length} words from ${xlsxPath}`);

    for (const row of jsonData) {
        const fileName = row["mp3"];
        const en = row["name"];
        const filePath = path.join(outputDir, fileName);

        if (await fileExists(filePath)) {
            console.log(`⏭️ Skipped (already exists): ${filePath}`);
            continue;
        }

        try {
            await synthesizeToFile(en, filePath);
            console.log(`✅ Created: ${filePath}`);

            // Avoid hitting the API rate limit
            await new Promise(r => setTimeout(r, 300));
        } catch (err) {
            console.error(`❌ Error creating audio for "${en}":`, err);
        }
    }

    console.log("🎯 All words processed.");
}

// Run the generator
try {
    await generateAudio();
} catch (err) {
    console.error('❌ Fatal error:', err);
    process.exit(1);
}