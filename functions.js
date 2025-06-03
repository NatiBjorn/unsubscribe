Office.initialize = () => {}

function extractUnsubscribeLink(body) {
  const unsubscribeRegex = /(https?:\/\/[\S]*?unsubscribe[\S]*)/gi;
  const matches = body.match(unsubscribeRegex);
  return matches ? matches[0] : null;
}

async function unsubscribeFromEmail(event) {
  Office.context.mailbox.item.body.getAsync("html", async (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const htmlBody = result.value;

      let link = extractUnsubscribeLink(htmlBody);

      if (!link) {
        try {
          const res = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
              'Authorization': 'Bearer YOUR_OPENAI_KEY',
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              model: "gpt-4",
              messages: [{
                role: "user",
                content: `Extract an unsubscribe URL from this email HTML:\n\n${htmlBody}`
              }]
            })
          });
          const data = await res.json();
          link = data.choices?.[0]?.message?.content?.match(/https?:\/\/\S+/)?.[0];
        } catch (e) {
          console.error("ChatGPT API error", e);
        }
      }

      if (link) {
        window.open(link, "_blank");
      } else {
        alert("No unsubscribe link found.");
      }
    } else {
      console.error("Failed to get email body:", result.error);
    }
    event.completed();
  });
}

window.unsubscribeFromEmail = unsubscribeFromEmail;
