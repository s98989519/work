async function processData() {
    const input = document.getElementById('userInput').value;
    const output = document.getElementById('output');
    
    output.innerText = "Processing... Please wait...";

    // 這裡換成你剛才在 Cloudflare 建立的 Worker 網址
    const PROXY_URL = 'https://你的worker名字.workers.dev/v1/chat/completions';
    
    // 再次提醒：Key 建議手動輸入或放在 Private Repo
    const API_KEY = '你的OpenAI金鑰'; 

    try {
        const response = await fetch(PROXY_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${API_KEY}`
            },
            body: JSON.stringify({
                model: "gpt-4o", 
                messages: [{role: "user", content: input}]
            })
        });

        const data = await response.json();
        output.innerText = data.choices[0].message.content;
    } catch (error) {
        output.innerText = "Error: System response timeout. Check connectivity.";
        console.error(error);
    }
}
