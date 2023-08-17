require("dotenv").config();
const ExcelJS = require('exceljs');
const { Configuration, OpenAIApi }= require("openai");

const configuration = new Configuration({
    apiKey: process.env.OPENAI_API_KEY,
});

const openai = new OpenAIApi(configuration);

const generatePromptA = (age,numberOfLines) => (`Write ${numberOfLines} important things famous people did when they were ${age} years old. Refer to a specific famous person in each sentence. State the age of ${age} in each sentence, different variations. Write all 200 facts in sequence and number them. Don't write an opening sentence, get straight to the point. Do not write a closing sentence.`)

const generatePromptB = (age,numberOfLines) => (`Write ${age} interesting and reliable facts that are different from each other, related to the number ${age}. The ${numberOfLines} facts can be from the fields of food and recipes, biology, zoology, politics, Different cultures, anthropology, Judaism, Islam, architecture, the sea, plants and botany, astronomy, space, the human body, entertainment, numerology, Musical instruments, sports, Guinness World Records, music, chemistry, art, literature, poetry, fantasy literature, famous people, philosophy, movies, technology, the New Age, geography, ecology, anthropology, linguistics, history, archeology, prehistory, space The outdoors, the military, dinosaurs, physics, languages, painting, brands, board games, medicine,Nutritional values of foods, vegetables and fruits, commerce, the economy, laws, dinosaurs, morality, computers, psychology, magic, fantasy, And in addition facts from any other field you can think of. Write 200 sentences in total, no less. Don't write an opening sentence, get straight to the point. Do not write a closing sentence.`)

const numberOfLines = 200;

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

const triggerGPT = async (prompt) => {
    const response = await openai.createChatCompletion({
        model: process.env.OPENAI_MODEL,
        messages: [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", content: prompt}
        ],
    });
    const content = response.data.choices[0].message.content;
    const lines = content.split('\n');
    return lines;
}

async function generateExcelForAges() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sentences');
    worksheet.addRow(['Number', 'Description', 'Age']);
    
    const allData = [];

    try {
        for (let age = 1; age <= 110; age++) {
            const contentForCurrentAge = generatePromptA(age, numberOfLines);
            const lines = await triggerGPT(contentForCurrentAge);
            const ageData = lines.map((line, index) => {
                const description = line.slice(line.indexOf(' ') + 1);
                return [index + 1, description, age];
            });

            allData.push(...ageData, ['-', '-', '-']);
            await delay(10000); // Prevent OpenAI API from throttling
        }
    } catch (error) {
        console.error("An error occurred:", error);
    } finally {
        allData.forEach(row => {
            worksheet.addRow(row);
        });
        await workbook.xlsx.writeFile('AgesAndPersonalities.xlsx');
    }
}

generateExcelForAges();