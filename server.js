require('dotenv').config();
const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const client = new Anthropic();

app.post('/api/suggest', async (req, res) => {
  const { workContent, weather, workers } = req.body;

  if (!workContent || !workContent.trim()) {
    return res.status(400).json({ error: '作業内容を入力してください' });
  }

  const prompt = `造園・公園管理の現場でのKY（危険予知）活動の内容を考えてください。

作業内容: ${workContent.trim()}
天候: ${(weather || '晴れ').trim()}
作業員数: ${workers ? workers + '名' : '不明'}

以下のJSON形式のみで返してください（説明文は不要）:
{
  "danger1": "危険のポイント①（具体的な危険の内容）",
  "action1": "私達はこうする①（具体的な対策）",
  "danger2": "危険のポイント②（具体的な危険の内容）",
  "action2": "私達はこうする②（具体的な対策）",
  "safetyGoal": "本日の安全目標（短く力強い一文）"
}`;

  try {
    const message = await client.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1024,
      messages: [
        { role: 'user', content: prompt }
      ]
    });

    const text = message.content[0].text.trim();
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return res.status(500).json({ error: 'AI応答の解析に失敗しました。再試行してください。' });
    }

    const result = JSON.parse(jsonMatch[0]);
    res.json(result);
  } catch (error) {
    console.error('API Error:', error.message);
    res.status(500).json({ error: error.message });
  }
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`KY作成AIツール起動中: http://localhost:${PORT}`);
});
