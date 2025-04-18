import axios from "axios";

const apiUrl = "";
const apiKey = "";
const deploymentName = "";

export const callService = async (query) => {
  var response = {};
  const config = {
    headers: {
      "api-key": apiKey,
      "Content-Type": "application/json",
    },
    params: {
      "api-version": "2023-07-01-preview",
    },
  };

  const prompt = {
    messages: [
      {
        role: "system",
        content: "You are a helpful assistant.",
      },
      {
        role: "user",
        content: query,
      },
    ],
    temperature: 1.0,
    n: 1,
  };

  const suggestion_prompt = {
    messages: [
      {
        role: "system",
        content: "You are a helpful assistant.",
      },
      {
        role: "user",
        content:
          "You are a prompt engineer tasked with generating follow-up prompts that delve deeper into the user's initial query. Given the following user prompt, provide a list of 3 follow-up prompts that could be used to refine the user's request and guide them towards a more specific or informative answer:"+query,
      },
    ],
    temperature: 1.0,
    n: 1,
  };

  const res = await axios.post(apiUrl + "openai/deployments/" + deploymentName + "/chat/completions", prompt, config);

  if (res.status == 200 && res.data != null) {
    const content = res.data.choices[0].message.content;
    const suggestion_res = await axios.post(
      apiUrl + "openai/deployments/" + deploymentName + "/chat/completions",
      suggestion_prompt,
      config
    );

    response["content"] = content;
    if (suggestion_res.status == 200 && suggestion_res.data != null) {
      const suggestion = suggestion_res.data.choices[0].message.content;
      console.log(suggestion);
      response["suggestions"] = suggestion;
    } else {
      throw Error(suggestion_res.data.error);
    }
  } else {
    throw Error(res.data.error);
  }
  return response;
};
