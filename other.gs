// スクリプトプロパティの取得用のヘルパー関数
function getScriptProperties() {
  const scriptProperties = {
    DATABASE_ID: PropertiesService.getScriptProperties().getProperty('DATABASE_ID'),
    NOTION_API_KEY: PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY'),
    OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY')
  };
  
  // 必要な定数が設定されているか確認
  const missingProperties = Object.entries(scriptProperties)
    .filter(([key, value]) => !value)
    .map(([key]) => key);
  
  if (missingProperties.length > 0) {
    throw new Error(`Missing required properties: ${missingProperties.join(', ')}`);
  }
  
  return scriptProperties;
}

function run() {
  const createLastTime = getLastCheckedTime();
  getPagesFromNotion(createLastTime);
}

function getPagesFromNotion(lastCheckedTime) {
  const properties = getScriptProperties();
  const url = `https://api.notion.com/v1/databases/${properties.DATABASE_ID}/query`;
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
    // console.log('Response Status:', response.getResponseCode());
    // console.log('Response Headers:', response.getAllHeaders());
    
    const responseText = response.getContentText();
    // console.log('Raw Response:', responseText);
    
    const data = JSON.parse(responseText);
    // console.log('Parsed Data:', JSON.stringify(data, null, 2));
    
    if (!data || !data.results) {
      throw new Error('Invalid response structure: missing results array');
    }
    if (data.results.length > 0) {
      data.results.sort((a, b) => {
        return new Date(a.properties['Created time'].created_time) - new Date(b.properties['Created time'].created_time);
      });

      data.results.forEach(page => {
        let pageId = page.id;
        let text = page.properties['単語'].title[0].text.content;
        let created_time = page.properties['Created time'].created_time;
        
        let wordPrompt = '「' + text + '」というエスペラント語を日本語に翻訳してください。翻訳した後は、必ず日本語の意味だけを記述してください。';
        let translatedWord = generateExampleSentence(text, wordPrompt);

        const translatedWordProperties = {
          '意味': {
            'rich_text': [
              {
                'text': {
                  'content': translatedWord
                }
              }
            ]
          }
        };
        updateNotionPageProperties(pageId, translatedWordProperties);

        let prompt = '「' + text + '」という単語を使って下記のフォーマットで簡単で短めのエスペラント語の例文とその後ろに()でその日本語翻訳を１つだけ作成してください。必ず１つだけにしてください。';
        let example = generateExampleSentence(text, prompt);

        const exampleText = {
          '例文': {
            'rich_text': [
              {
                'text': {
                  'content': example
                }
              }
            ]
          }
        };
        updateNotionPageProperties(pageId, exampleText);

        let etymology = '「' + text + '」という単語の語源を解説して。解説はできるだけ短く50語以内に解説して;';
        let etymologyText = generateExampleSentence(text, etymology);

        const etymologyInput = {
          '語源': {
            'rich_text': [
              {
                'text': {
                  'content': etymologyText
                }
              }
            ]
          }
        };
        updateNotionPageProperties(pageId, etymologyInput);

        let pronunciation = '「' + text + '」という単語のエスペラント語のIPAの発音記号だけを書いてください。必ず発音記号だけを書いてください。';
        let pronunciationSymbol = generateExampleSentence(text, pronunciation);

        const pronunciationInput = {
          '発音': {
            'rich_text': [
              {
                'text': {
                  'content': pronunciationSymbol
                }
              }
            ]
          }
        };
        updateNotionPageProperties(pageId, pronunciationInput);

        saveLastCheckedTime(created_time);
      });
    } else {
      console.log("データは存在しませんでした");
    }
  } catch (error) {
    console.error('Detailed error:', {
      message: error.message,
      stack: error.stack,
      responseData: response ? responseText : 'No response'
    });
    throw error;
  }
}

function updateNotionPageProperties(pageId, properties) {
  const scriptProperties = getScriptProperties();
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

function generateExampleSentence(word, prompt) {
  const properties = getScriptProperties();
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

function getLastCheckedTime() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastCheckedTime = sheet.getRange('A1').getValue();
  // 値が空白や無効ならデフォルトの日時を設定
  if (!lastCheckedTime || isNaN(new Date(lastCheckedTime).getTime())) {
    return "2024-01-01T00:00:00.000Z"; // デフォルト値
  }
  return lastCheckedTime;
}

function saveLastCheckedTime(time) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange('A1').setValue(time);
}
