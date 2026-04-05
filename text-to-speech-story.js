import fs from "fs";
import path from "path";
import * as XLSX from 'xlsx';
import * as sdk from "microsoft-cognitiveservices-speech-sdk";
import 'dotenv/config';
import process from 'process';

//const xlsxPath = 'xlsx/276_Question.xlsx';
//const outputDir = 'mp3/question';

//const xlsxPath = 'xlsx/828_Answer.xlsx';
//const outputDir = 'mp3/answer';

const xlsxPath = 'xlsx/Story.xlsx';
//const outputDir = 'mp3/story/male';
const outputDir = 'mp3/story/female';

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
        //speechConfig.speechSynthesisVoiceName = "en-US-AdamMultilingualNeural"; 
        speechConfig.speechSynthesisVoiceName = "en-US-NovaTurboMultilingualNeural"; 
         
        // en-US-JennyNeural
        // en-US-RyanMultilingualNeural
        // en-US-BrianMultilingualNeural
        // en-US-Brian:DragonHDLatestNeural
        // en-US-Ava:DragonHDLatestNeural
        //
        // Dragon HD voices
        // Male
        // en-US-AdamMultilingualNeural
        // en-US-AndrewMultilingualNeural
        // Female
        // en-US-NovaTurboMultilingualNeural
        // en-US-AvaMultilingualNeural
        // en-US-AmandaMultilingualNeural

        speechConfig.speechSynthesisOutputFormat =
            sdk.SpeechSynthesisOutputFormat.Audio44Khz128KBitRateMonoMp3;

        const audioConfig = sdk.AudioConfig.fromAudioFileOutput(filePath);
        const synthesizer = new sdk.SpeechSynthesizer(speechConfig, audioConfig);

        const timeoutId = setTimeout(() => {
            synthesizer.close();
            reject(new Error(`Synthesis timed out for: ${text}`));
        }, 30_000);

        synthesizer.speakTextAsync(
            text,
            result => {
                clearTimeout(timeoutId);
                synthesizer.close();
                if (result.reason === sdk.ResultReason.SynthesizingAudioCompleted) {
                    resolve();
                } else {
                    reject(new Error(result.errorDetails ?? `Synthesis failed: ${result.reason}`));
                }
            },
            err => {
                clearTimeout(timeoutId);
                synthesizer.close();
                reject(typeof err === 'string' ? new Error(err) : err);
            }
        );
    });
}

// Generate audio for all stories sequentially (basic, intermediate, advanced per row)
async function generateAudio() {

    const workbook = XLSX.read(fs.readFileSync(xlsxPath));
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    console.log(`📄 Loaded ${jsonData.length} rows from ${xlsxPath}`);

    const levels = [
        { textCol: 'basic',        mp3Col: 'basic_mp3' },
        { textCol: 'intermediate', mp3Col: 'intermediate_mp3' },
        { textCol: 'advanced',     mp3Col: 'advanced_mp3' },
    ];

    for (const row of jsonData) {
        for (const { textCol, mp3Col } of levels) {
            const text = row[textCol];
            const fileName = row[mp3Col];

            if (!text || !fileName) {
                console.warn(`⚠️  Skipping row id=${row["id"]} level="${textCol}": missing text or filename`);
                continue;
            }

            const filePath = path.join(outputDir, fileName.replace(/[ ,]+/g, '_'));

            if (await fileExists(filePath)) {
                console.log(`⏭️ Skipped (already exists): ${filePath}`);
                continue;
            }

            const maxRetries = 3;
            let success = false;
            for (let attempt = 1; attempt <= maxRetries; attempt++) {
                try {
                    await synthesizeToFile(text, filePath);
                    console.log(`✅ Created: ${filePath}`);
                    success = true;
                    // Avoid hitting the API rate limit
                    await new Promise(r => setTimeout(r, 300));
                    break;
                } catch (err) {
                    console.warn(`⚠️  Attempt ${attempt}/${maxRetries} failed for "${textCol}" (id=${row["id"]}): ${err.message}`);
                    if (attempt < maxRetries) {
                        const backoff = attempt * 2000;
                        console.log(`   Retrying in ${backoff / 1000}s...`);
                        await new Promise(r => setTimeout(r, backoff));
                    }
                }
            }
            if (!success) {
                console.error(`❌ Giving up on "${textCol}" (id=${row["id"]}) after ${maxRetries} attempts.`);
            }
        }
    }

    console.log("🎯 All stories processed.");
}

// Run the generator
try {
    await generateAudio();
    process.exit(0);
} catch (err) {
    console.error('❌ Fatal error:', err);
    process.exit(1);
}