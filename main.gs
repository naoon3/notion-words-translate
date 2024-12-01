// スクリプトプロパティの取得用のヘルパー関数
function getScriptProperties(language) {
  const languageConfig = LANGUAGE_CONFIGS[language];
  const scriptProperties = {
    DATABASE_ID: languageConfig.databaseId,
    NOTION_API_KEY: PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY'),
    DEEPL_API_KEY: PropertiesService.getScriptProperties().getProperty('DEEPL_API_KEY'),
    OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY')
  };
  
  // 必要な定数が設定されているか確認
  const missingProperties = Object.entries(scriptProperties)
    .filter(([key, value]) => !value)
    .map(([key]) => key);
  
  if (missingProperties.length > 0) {
    throw new Error(`Missing required properties for ${language}: ${missingProperties.join(', ')}`);
  }
  
  return scriptProperties;
}

// 言語設定を追加
const LANGUAGE_CONFIGS = {
  'English': {
    databaseId: '', // Notion DBのここに入力
    targetLang: 'JA',
    translationLang: 'JA',
    wordProperty: '単語',
    meaningProperty: '意味',
    exampleProperty: '例文',
    etymologyProperty: '語源',
    pronunciationProperty: '発音',
    pronunciationPrompt: (word) => `"${word}"という単語の英語の発音記号だけを書いてください。発音の基準はアメリカ英語でお願いします。必ずIPAの発音記号だけを書いてください。`,
    examplePrompt: (word) => `「${word}」という単語を使って下記のフォーマットで簡単で短めの英語の例文とその後ろに()でその日本語翻訳を１つだけ作成してください。必ず１つだけにしてください。`,
    etymologyPrompt: (word) => `「${word}」という単語の語源を解説して。解説はできるだけ短く50語以内に解説して;`
  },
  'Chinese': {
    databaseId: '', // Notion DBのここに入力
    targetLang: 'JA',
    translationLang: 'JA',
    wordProperty: '単語',
    meaningProperty: '意味',
    exampleProperty: '例文',
    etymologyProperty: '語源',
    pronunciationProperty: '発音',
    pronunciationPrompt: (word) => `"${word}"という中国語の単語の発音をピンインで書いてください。必ず中国語のピンインだけを書いてください。`,
    examplePrompt: (word) => `「${word}」という単語を使って下記のフォーマットで簡単で短めの中国語の例文とその後ろに()でその日本語翻訳を１つだけ作成してください。必ず１つだけにしてください。`,
    etymologyPrompt: (word) => `「${word}」という単語の語源を解説して。解説はできるだけ短く50語以内に解説して;`
  },
  'Russian': {
    databaseId: '', // Notion DBのここに入力
    targetLang: 'JA',
    translationLang: 'JA',
    wordProperty: '単語',
    meaningProperty: '意味',
    exampleProperty: '例文',
    etymologyProperty: '語源',
    pronunciationProperty: '発音',
    pronunciationPrompt: (word) => `"${word}"というロシア語の単語のキリル文字での発音記号を書いてください。必ずIPAの発音記号だけを書いてください。`,
    examplePrompt: (word) => `「${word}」という単語を使って下記のフォーマットで簡単で短めのロシア語の例文とその後ろに()でその日本語翻訳を１つだけ作成してください。必ず１つだけにしてください。`,
    etymologyPrompt: (word) => `「${word}」という単語の語源を解説して。解説はできるだけ短く50語以内に解説して;`
  },
  'Korean': {
    databaseId: '', // Notion DBのここに入力
    targetLang: 'JA',
    translationLang: 'JA',
    wordProperty: '単語',
    meaningProperty: '意味',
    exampleProperty: '例文',
    etymologyProperty: '語源',
    pronunciationProperty: '発音',
    pronunciationPrompt: (word) => `"${word}"という韓国語の単語の発音記号を書いてください。必ずIPAの発音記号だけを書いてください。`,
    examplePrompt: (word) => `「${word}」という単語を使って下記のフォーマットで簡単で短めの韓国語の例文とその後ろに()でその日本語翻訳を１つだけ作成してください。必ず１つだけにしてください。`,
    etymologyPrompt: (word) => `「${word}」という単語の語源を解説して。解説はできるだけ短く50語以内に解説して;`
  }
};

function runForLanguages() {
  const languages = ["English", "Chinese", "Russian", "Korean"];
  for (const language of languages) {
    run(language);
  }
}

function run(language) {
  const createLastTime = getLastCheckedTime(language);
  getPagesFromNotion(createLastTime, language);
}

function getPagesFromNotion(lastCheckedTime, language) {
  const languageConfig = LANGUAGE_CONFIGS[language];
  const properties = getScriptProperties(language);
  
  // 言語別のデータベースIDを使用
  const url = `https://api.notion.com/v1/databases/${languageConfig.databaseId}/query`;
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${properties.NOTION_API_KEY}`,
      'Content-Type': 'application/json',
      'Notion-Version': "2022-06-28"
    },
    payload: JSON.stringify({
      filter: {
        property: "Created time",
        created_time: {
          after: lastCheckedTime
        }
      },
      sorts: [
        {
          property: "Created time",
          direction: "ascending"
        }
      ]
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.results.length > 0) {
      data.results.forEach(page => {
        let pageId = page.id;
        let text = page.properties[languageConfig.wordProperty].title[0].text.content;
        let translatedText = translateTextWithDeepL(text, languageConfig.targetLang, language);

        const meaningProperties = {
          [languageConfig.meaningProperty]: {
            'rich_text': [
              {
                'text': {
                  'content': translatedText
                }
              }
            ]
          }
        };

        updateNotionPageProperties(pageId, meaningProperties, language);

        let examplePrompt = languageConfig.examplePrompt(text);
        let example = generateExampleSentence(text, examplePrompt, language);

        const exampleProperties = {
          [languageConfig.exampleProperty]: {
            'rich_text': [
              {
                'text': {
                  'content': example
                }
              }
            ]
          }
        };
        updateNotionPageProperties(pageId, exampleProperties, language);

        let etymologyPrompt = languageConfig.etymologyPrompt(text);
        let etymologyText = generateExampleSentence(text, etymologyPrompt, language);

        const etymologyProperties = {
          [languageConfig.etymologyProperty]: {
            'rich_text': [
              {
                'text': {
                  'content': etymologyText
                }
              }
            ]
          }
        };
        updateNotionPageProperties(pageId, etymologyProperties, language);

        let pronunciationPrompt = languageConfig.pronunciationPrompt(text);
        let pronunciationSymbol = generateExampleSentence(text, pronunciationPrompt, language);

        const pronunciationProperties = {
          [languageConfig.pronunciationProperty]: {
            'rich_text': [
              {
                'text': {
                  'content': pronunciationSymbol
                }
              }
            ]
          }
        };
        updateNotionPageProperties(pageId, pronunciationProperties, language);

        saveLastCheckedTime(page.properties['Created time'].created_time, language);
      });
    } else {
      console.log("データは存在しませんでした");
    }
  } catch (error) {
    console.error('Detailed error:', {
      message: error.message,
      stack: error.stack,
      language: language
    });
    throw error;
  }
}

function translateTextWithDeepL(text, targetLang = 'JA', language) {
  const properties = getScriptProperties(language);
  const url = `https://api-free.deepl.com/v2/translate?auth_key=${properties.DEEPL_API_KEY}`;
  const payload = {
    'text': text,
    'target_lang': targetLang
  };

  const options = {
    'method': 'post',
    'payload': payload,
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse.translations && jsonResponse.translations.length > 0) {
      return jsonResponse.translations[0].text;
    }
  } catch (error) {
    console.error('Error during translation with DeepL:', error);
    return text; // 翻訳に失敗した場合は元のテキストを返す
  }
}

function updateNotionPageProperties(pageId, properties, language) {
  const scriptProperties = getScriptProperties(language);
  const url = `https://api.notion.com/v1/pages/${pageId}`;

  const options = {
    method: 'patch',
    headers: {
      'Authorization': `Bearer ${scriptProperties.NOTION_API_KEY}`,
      'Content-Type': 'application/json',
      'Notion-Version': "2021-08-16"
    },
    payload: JSON.stringify({
      properties: properties
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const jsonResponse = JSON.parse(response.getContentText());
  } catch (error) {
    console.error('Error updating Notion page:', error);
  }
}

function generateExampleSentence(word, prompt, language) {
  const properties = getScriptProperties(language);
  const apiUrl = 'https://api.openai.com/v1/chat/completions';

  const messages = [
    { 'role': 'system', 'content': prompt },
    { 'role': 'user', 'content': prompt },
  ];

  const headers = {
    'Authorization': 'Bearer ' + properties.OPENAI_API_KEY,
    'Content-type': 'application/json',
    'X-Slack-No-Retry': 1
  };

  const options = {
    'muteHttpExceptions': true,
    'headers': headers,
    'method': 'POST',
    'payload': JSON.stringify({
      'model': 'gpt-3.5-turbo',
      'max_tokens': 2000,
      'temperature': 0.9,
      'messages': messages
    })
  };

  const response = JSON.parse(UrlFetchApp.fetch(apiUrl, options).getContentText());
  return response.choices[0].message.content;
}

function getLastCheckedTime(language) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${language}LastChecked`);
  if (!sheet) {
    throw new Error(`Sheet for ${language} not found`);
  }
  const lastCheckedTime = sheet.getRange('A1').getValue();
  // 値が空白や無効ならデフォルトの日時を設定
  if (!lastCheckedTime || isNaN(new Date(lastCheckedTime).getTime())) {
    return "2024-01-01T00:00:00.000Z"; // デフォルト値
  }
  return lastCheckedTime;
}

function saveLastCheckedTime(time, language) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${language}LastChecked`);
  if (!sheet) {
    throw new Error(`Sheet for ${language} not found`);
  }
  sheet.getRange('A1').setValue(time);
}
