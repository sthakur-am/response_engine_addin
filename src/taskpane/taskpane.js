/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import axios from "axios";
import { v4 as uuidv4 } from "uuid";
import services from "./services.json";
import Showdown from "showdown";

// OfficeExtension.config.extendedErrorLogging = true;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (Office.context.requirements.isSetSupported("WordApi", "1.9")) {
      // Your code that uses Word JavaScript API 1.9 features
      console.log("Word JavaScript API 1.9 is supported.");
    } else {
      console.log("Word JavaScript API 1.9 is not supported.");
    }

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("app").style.display = "none";
    document.getElementById("upload-dialog").classList.add("hidden");
    document.getElementById("completion-page").classList.add("hidden");

    // document.getElementById("apply-template").onclick = loadTemplateFile;

    document.getElementById("generateTemplate").onclick = showGenerateTemplateDialog;
    document.getElementById("close-dialog").onclick = hideGenerateTemplateDialog;
    document.getElementById("cancel-button").onclick = hideGenerateTemplateDialog;
    document.getElementById("complete-drafting").onclick = toggleCompleteDraftingButton;

    const apiUrls = services.apiUrls;
    console.log(apiUrls);

    renderAddIn();
  }
});

var converter = new Showdown.Converter();

export const toggleCompleteDraftingButton = () => {
  const docid = sessionStorage.getItem("docid");
  complete_rfp_draft(docid);
};

var enabled = false;

export const showGenerateTemplateDialog = () => {
  console.log("Hi from generateTemplate");
  document.getElementById("landing-page").classList.add("hidden");

  document.getElementById("upload-dialog").classList.remove("hidden");
  document.getElementById("upload-dialog").classList.add("open");
  initializeUploadDialog();
};

export const hideGenerateTemplateDialog = () => {
  document.getElementById("upload-dialog").classList.remove("open");
  document.getElementById("landing-page").classList.remove("hidden");
};

export async function renderAddIn() {
  return Word.run(async (context) => {
    document.getElementById("app").classList.add("hidden");
    showSpinner();
    toggleDiv();
    const sections = context.document.sections;
    context.load(sections);

    await context.sync();
    const header = sections.items[0].getHeader("primary");
    context.load(header, "text");
    await context.sync();

    if (header.text.trim() != "") {
      const docid = header.text.trim();
      console.log("Document ID: " + docid);
      document.getElementById("landing-page").classList.add("hidden");
      sessionStorage.setItem("docid", docid);

      // Check status of this document
      const res = await callDraftStstusService(docid);
      const draft = res["data"];
      console.log(draft);
      if (draft["status"] == 0) {
        // load uploaded documents
        const documents = await callCAIODocService(docid);
        if (documents != null && documents.length > 0) {
          console.log(documents);
          sessionStorage.setItem("documents", JSON.stringify(documents));
        }

        // load document tasks
        const res = await callCAIOTaskService(docid);
        if (res != null) {
          console.log(res.data);
          var tasks = null;
          try {
            tasks = JSON.parse(sessionStorage.getItem("tasks"));
          } catch (error) {
            console.log(error);
          }

          var new_tasks = res.data;
          if (tasks != null) {
            new_tasks = res.data;
            new_tasks[3].options = tasks[3].options;
            new_tasks[4].options = tasks[4].options;
          }
          sessionStorage.setItem("tasks", JSON.stringify(new_tasks));
          renderTasks(new_tasks);
          updateProgress(new_tasks);
          initializeChat();
          initializeDocumentsRepo();
          document.getElementById("app").classList.remove("hidden");
          document.getElementById("app").style.display = "block";
        }
      } else {
        console.log("This document is no longer available for drafting.");
        renderCompletionPage(draft);
      }
    } else {
      console.log("No DocId");
      const docid = uuidv4();
      sessionStorage.setItem("docid", docid);
      console.log("Document ID # " + docid);

      sessionStorage.removeItem("tasks");
      sessionStorage.removeItem("documents");

      initializeChat();
      initializeDocumentsRepo();
      document.getElementById("app").classList.add("hidden");
      document.getElementById("landing-page").classList.remove("hidden");
    }
    hideSpinner();
    toggleDiv();
  });
}

export async function renderCompletionPage(doc) {
  console.log("Rendering completion page");

  // render duration
  const startDate = doc["start_date"] ? doc["start_date"] : "NA";
  const endDate = doc["end_date"] ? doc["end_date"] : "NA";

  const metadata = document.getElementById("completion-meta");
  metadata.innerHTML = `
    <div>
      <h3 class="ms-font-l">Started</h3>
      <p class="text-lg font-light">${startDate}</p>
    </div>
    <div>
      <h3 class="ms-font-l">Completed</h3>
      <p class="text-lg font-light">${endDate}</p>
    </div>
  `;

  // render documents
  const documentsList = document.getElementById("completion-documents-list");
  if (sessionStorage.getItem("documents") != null) {
    var documents = JSON.parse(sessionStorage.getItem("documents"));
    console.log(documents);

    if (documents.length > 0) {
      documentsList.innerHTML = "";

      documents.forEach((doc) => {
        console.log(doc);
        const docElement = createDocumentElement(doc);
        documentsList.appendChild(docElement);
      });
    } else {
      documentsList.innerHTML = `
      <div class="empty-state">
        <p>No documents found for this draft.</p>
      </div>
      `;
    }
  } else {
    documentsList.innerHTML = `
      <div class="empty-state">
        <p>No documents found for this draft.</p>
      </div>
      `;
  }

  document.getElementById("app").classList.add("hidden");
  document.getElementById("completion-page").classList.remove("hidden");
}

export const toggleDiv = () => {
  var status = !enabled;
  enabled = status;
};

export const callCAIOChatService = async (docid, message) => {
  const functionUrl = services.apiUrls.query_index;
  const data = {
    doc_id: docid,
    query: message,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };
  console.log("Calling response engine chat service");
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
    }else {
      console.log("No result");
    }
    return res.data.data;
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callExtractRFPDataService = async (docid) => {
  const functionUrl = services.apiUrls.get_rfp_metadata;
  const data = {
    doc_id: docid,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };
  console.log("Calling response engine document extract service");
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
      JSON.parse(res.data.data);
    }else {
      console.log("No result");
      return null;
    }
    return res.data.data;
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callDocumentSummaryService = async (docid, documentPath) => {
  const functionUrl = services.apiUrls.summarize_doc;
  const data = {
    doc_id: docid,
    doc_path: documentPath,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };
  console.log("Calling response engine document summary service");
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
    }else {
      console.log("No result");
      return null;
    }
    return res.data.data;
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callCAIODocService = async (docid) => {
  const functionUrl = services.apiUrls.list_docs;
  const data = {
    doc_id: docid,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };
  console.log("Calling response engine doc service");
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
    } else {
      console.log("No result");
      return null;
    }
    return res.data.data;
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callCAIODocAsBase64Service = async (docid) => {
  const functionUrl = services.apiUrls.get_doc_as_base64;
  const data = {
    doc_id: docid,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };
  console.log("Calling response engine doc as base64 service");
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
    } else {
      console.log("No result");
      return null;
    }
    return res.data.data;
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callCAIOSummaryService = async (docId) => {
  const functionUrl = services.apiUrls.list_rfp_asks;
  const data = {
    doc_id: docId,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };
  console.log("Calling response engine summarize rfp service");
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
      return res.data;
    } else {
      console.log("No result");
      return null;
    }
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callCAIOService = async (docId) => {
  const functionUrl = services.apiUrls.generate_response_sections;
  const data = {
    doc_id: docId,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };
  console.log("Calling response engine generate structure service");
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
      const outline = res.data.data;
      sessionStorage.setItem("outline", JSON.stringify(outline));
      console.log(outline);
      return res.data;
    }else {
      console.log("No result");
      return null;
    }
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callCAIOContentService = async (docId, sectionHeading, sectionRequirement) => {
  const functionUrl = services.apiUrls.generate_section_content;
  const data = {
    doc_id: docId,
    section_title: sectionHeading,
    section_requirements: sectionRequirement,
    // rfp_text: rfpText,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };
  console.log("Calling response engine generate content service");
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
      const outline = res.data.data;
      console.log(outline);
      return res.data;
    } else {
      console.log("No result");
      return null;
    }
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callCAIOReviewService = async (docId, content) => {
  const functionUrl = services.apiUrls.review_section_content;
  const data = {
    doc_id: docId,
    content: content,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };
  console.log("Calling response engine review content service");
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
      const outline = res.data.data;
      console.log(outline);
      return res.data;
    } else {
      console.log("No result");
      return null;
    }
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callCAIOTaskService = async (docid) => {
  const functionUrl = services.apiUrls.list_rfp_tasks;
  const formData = new FormData();
  formData.append("doc_id", docid);

  const config = {
    headers: {
      "Content-Type": "multipart/form-data",
    },
  };
  console.log("Calling response engine generate task service");
  try {
    const res = await axios.post(functionUrl, formData, config);
    console.log(res);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
      const tasklist = res.data.data;
      sessionStorage.setItem("tasklist", JSON.stringify(tasklist));
      console.log(tasklist);
      return res.data;
    } else {
      console.log("No result");
      return null;
    }
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callCAIOTaskUpdateService = async (docid, task_status) => {
  const functionUrl = services.apiUrls.update_rfp_task;
  const data = {
    doc_id: docid,
    task_status: [task_status],
  };
  console.log(data);
  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };
  console.log("Calling response engine task update service");
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
      const tasklist = res.data.data;
      sessionStorage.setItem("tasklist", JSON.stringify(tasklist));
      console.log(tasklist);
      return res.data;
    } else {
      console.log("No result");
      return null;
    }
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callUploadFilesService = async (docid, items) => {
  const functionUrl = services.apiUrls.upload_files;
  const formData = new FormData();
  for (let i = 0; i < items.length; i++) {
    var item = items[i];
    if (item != null) {
      console.log(item["type"]);
      console.log(item["file"].name);
      formData.append(item["type"] + ":" + item["file"].name, item["file"]);
    }
  }
  formData.append("doc_id", docid);

  const config = {
    headers: {
      "Content-Type": "multipart/form-data",
    },
  };
  console.log("Calling response engine generate upload files service for " + docid);
  try {
    const res = await axios.post(functionUrl, formData, config);
    console.log(res);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
      const filelist = res.data.data;
      console.log(filelist);
      return res.data;
    } else {
      console.log("Error :" + res.data.data);
      console.log("No result");
      return null;
    }
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callDraftStstusService = async (docid) => {
  const functionUrl = services.apiUrls.drafting_status;
  const data = {
    doc_id: docid,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };

  console.log("Calling response engine drafting status service for " + docid);
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
      return res.data;
    } else {
      console.log("Error :" + res.data.data);
      console.log("No result");
      return null;
    }
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export const callCompleteDraftingService = async (docid) => {
  const functionUrl = services.apiUrls.complete_drafting;
  const data = {
    doc_id: docid,
  };

  const config = {
    headers: {
      "Content-Type": "application/json",
    },
  };

  console.log("Calling response engine complete drafting service for " + docid);
  try {
    const res = await axios.post(functionUrl, data, config);
    console.log(res);
    console.log(res.data);
    if (res.data.status == "success" && res.data.data != null) {
      console.log(res.data.data);
      return res.data;
    } else {
      console.log("Error :" + res.data.data);
      console.log("No result");
      return null;
    }
  } catch (error) {
    console.log(error);
    showStatusMessage(error, "error");
    return null;
  }
};

export async function insertDataTable(data) {
  return Word.run(async (context) => {
    const body = context.document.body;
    const tableParagraph = body.insertParagraph("RFP Metadata Table", Word.InsertLocation.end);
    tableParagraph.styleBuiltIn = Word.Style.heading2;
    const tableData = [];
    tableData.push(["Metadata", "Value"]);
    console.log(data);
    var count = 0;
    for (const key in data) {
      const value = data[key];
      count = count + 1;
      console.log(key);
      console.log(value);
      tableData.push([key, value]);
    }
    console.log(tableData);
    const table = tableParagraph.insertTable(count + 1, 2, Word.InsertLocation.after, tableData);
    table.styleBuiltIn = Word.Style.normal;
    await context.sync();
  });
}

export async function insertSection(heading, text) {
  return Word.run(async (context) => {
    const body = context.document.body;

    if (text != null) {
      // Add title
      const title = body.insertParagraph(heading, Word.InsertLocation.end);
      title.styleBuiltIn = Word.Style.heading2;

      // Add body text
      if (text.includes("\n\n")) {
        var paragraphs = text.split("\n\n");
        paragraphs.forEach(function (para, index) {
          const paragraph = body.insertParagraph(para, Word.InsertLocation.end);
          paragraph.styleBuiltIn = Word.Style.normal;
        });
      } else {
        const paragraph = body.insertParagraph(text, Word.InsertLocation.end);
        paragraph.styleBuiltIn = Word.Style.normal;
      }
    }
    await context.sync();
  });
}

export async function insertSummary(summary) {
  return Word.run(async (context) => {
    const body = context.document.body;
    console.log(summary);
    try {
      const summaryParagraph = body.insertParagraph("RFP Requirements", Word.InsertLocation.end);
      summaryParagraph.styleBuiltIn = Word.Style.heading2;

      const contentParagraph = summaryParagraph.insertParagraph("", Word.InsertLocation.after);
      contentParagraph.styleBuiltIn = Word.Style.normal;
      contentParagraph.insertHtml(converter.makeHtml(summary), Word.InsertLocation.end);
      // contentParagraph.styleBuiltIn = Word.Style.normal;
      // var chunks = summary.split("\n");
      // for (let i = 0; i < chunks.length; i++) {
      //   const range = contentParagraph.insertParagraph(chunks[i], Word.InsertLocation.before);
      //   range.styleBuiltIn = Word.Style.normal;
      // }
      await context.sync();
    } catch (error) {
      console.log(error);
    }
  });
}

export async function insertOutline(outline) {
  return Word.run(async (context) => {
    const body = context.document.body;
    console.log(outline);
    try {
      // const items = JSON.parse(outline);
      Object.keys(outline).forEach(function (key) {
        console.log(key);
        const section = outline[key];
        // const title = body.insertParagraph(outline[key], Word.InsertLocation.end);

        const title = body.insertParagraph(section["title"], Word.InsertLocation.end);
        title.styleBuiltIn = Word.Style.heading2;

        const sectionParagraph = body.insertParagraph("", Word.InsertLocation.end);
        // const sectionParagraph = body.insertParagraph(metadata, Word.InsertLocation.end);
        sectionParagraph.styleBuiltIn = Word.Style.normal;

        sectionParagraph.getRange(Word.RangeLocation.end).insertText("Description: " + section["description"], Word.InsertLocation.end);
        // // desc.styleBuiltIn = Word.Style.normal;
        sectionParagraph.getRange(Word.RangeLocation.end).insertBreak(Word.BreakType.line, Word.InsertLocation.after);
        sectionParagraph.getRange(Word.RangeLocation.end).insertText("RFP requirements: " + section["requirements"], Word.InsertLocation.end);
        sectionParagraph.getRange(Word.RangeLocation.end).insertBreak(Word.BreakType.line, Word.InsertLocation.after);
        sectionParagraph.getRange(Word.RangeLocation.end).insertText("Content Prompt: " + section["prompt"], Word.InsertLocation.end);
      });
      await context.sync();
    } catch (error) {
      console.log(error);
    }
  });
}

export async function insertSectionContent(sectionTitle, content) {
  return Word.run(async (context) => {
    const body = context.document.body;
    try {
      const paragraphs = body.paragraphs;
      context.load(paragraphs, "text, style");
      await context.sync();

      var sectionParagraph = null;
      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        if (paragraph.style === "heading 2" && paragraph.text.trim() === sectionTitle) {
          console.log("Found the paragraph");
          console.log(paragraph.text);
          console.log(paragraph.style);
          sectionParagraph = paragraph.getNext();

        }
      }

      context.load(sectionParagraph);
      await context.sync();

      // sectionParagraph.insertParagraph

      const contentParagraph = sectionParagraph.insertParagraph("", Word.InsertLocation.after);
      contentParagraph.styleBuiltIn = Word.Style.normal;
      // const contentParagraph = body.insertParagraph("", Word.InsertLocation.after);
      contentParagraph.insertHtml(converter.makeHtml(content), Word.InsertLocation.end);
      // contentParagraph.insertText(content, Word.InsertLocation.end);
      

      await context.sync()
    } catch (error) {
      console.log(error);
    }
  });
}

export async function run_summary() {
  // return Word.run(async (context) => {
  toggleDiv();

  var docid = sessionStorage.getItem("docid");
  const res = await callCAIOSummaryService(docid);
  if (res != null) {
    console.log(res.data);
    await insertSummary(res.data);
  }
  toggleDiv();
  //   await context.sync();
  // });
}

export async function run_outline() {
  return Word.run(async (context) => {
    toggleDiv();

    var docid = sessionStorage.getItem("docid");
    const res = await callCAIOService(docid);
    if (res != null) {
      console.log(res.data);
      insertOutline(res.data);
      var options = [];
      var counter = 0;
      const outline = res.data;
      Object.keys(outline).forEach(function (key) {
        let option = {};
        counter = counter + 1;
        console.log(key);
        option["id"] = counter;
        const section = outline[key];
        option["label"] = section["title"];
        option["selected"] = false;

        options.push(option);
      });
      showStatusMessage("Added content sections to the document");
      add_task_options(4, options);
    }
    toggleDiv();
    await context.sync();
  });
}

export async function run_content(sectionTitle) {
  return Word.run(async (context) => {
    toggleDiv();
    var docid = sessionStorage.getItem("docid");
    var title = sectionTitle;

    const paragraphs = context.document.body.paragraphs;
    context.load(paragraphs, "text, style");
    await context.sync();

    var prompt = null;
    var sectionParagraph = null;
    for (let i = 0; i < paragraphs.items.length; i++) {
      const paragraph = paragraphs.items[i];
      if (paragraph.style === "heading 2" && paragraph.text.trim() === sectionTitle) {
        console.log("Found the paragraph");

        sectionParagraph = paragraph.getNext();
      }
    }

    context.load(sectionParagraph, "text");
    await context.sync();

    const text = sectionParagraph.text.toLowerCase();
    console.log(text);
    if (text.includes("content prompt:")) {
      prompt = text.split("content prompt:")[1].trim()
    }
    
    console.log(prompt);

    if (prompt == null) {
      prompt = "Generate content for the section '" + title + "'";
    }

    if (text != "") {
      const res = await callCAIOContentService(docid, title, prompt);
      // const res = prompt;
      if (res != null) {
        console.log(res.data);
        // console.log(res.data);
        // var content = res.data.split("\n").join(" ");
        // console.log(res);
        await insertSectionContent(title, res.data);

        // Add option to review task item for this content
        var options = [];
        var option = {};
        // option["id"] = counter;
        option["label"] = sectionTitle;
        option["selected"] = false;
        options.push(option);

        add_task_options(5, options);
      }
      showStatusMessage("Added content for section " + sectionTitle);
      toggleDiv();
    } else {
      console.log("No content found");
      showStatusMessage("Error adding content for section " + sectionTitle, "error");
      toggleDiv();
    }

    // await context.sync();
  });
}

export async function run_review(sectionTitle) {
  return Word.run(async (context) => {
    toggleDiv();
    console.log("running review");
    const paragraphs = context.document.body.paragraphs;
    context.load(paragraphs, "text, style");
    await context.sync();

    var docid = sessionStorage.getItem("docid");
    var contentParagraph = null;
    for (let i = 0; i < paragraphs.items.length; i++) {
      const paragraph = paragraphs.items[i];
      if (paragraph.style === "heading 2" && paragraph.text.trim() === sectionTitle) {
        console.log("Found the paragraph");
        contentParagraph = paragraph.getNext().getNext();
      }
    }
    var hasContentParagraphs = true;
    while (hasContentParagraphs) {
      contentParagraph = contentParagraph.getNext();
      hasContentParagraphs = await review_section_paragraph(context, docid, contentParagraph)
    }

    // context.load(contentParagraph, "text");
    // await context.sync();
    // text = contentParagraph.text;
    // // print("The section text");
    // // console.log(text);

    // var docid = sessionStorage.getItem("docid");
    // if (text != "") {
    //   console.log(text);
    //   const res = await callCAIOReviewService(docid, text);
    //   if (res != null) {
    //     console.log(res.data);
    //     var content = res.data.split("\n").join(" ");
    //     console.log(content);
    //     insertSectionContent(sectionTitle, content);
    //     showStatusMessage("Added reviewed content for the section " + sectionTitle);
    //   }
    //   toggleDiv();
    // } else {
    //   console.log("No content found");
    //   toggleDiv();
    //   showStatusMessage("Error finding content for the section " + sectionTitle);
    // }

    // await context.sync();
  });
}

export async function review_section_paragraph(context, docid, contentParagraph) {
  context.load(contentParagraph, "text, style");
  await context.sync();

  if (contentParagraph.style === "heading 2") {
    return false;
  }

  var text = contentParagraph.text;
  console.log(text);

  if (text.indexOf(".") > 0) {
    console.log("Found the paragraph to review")
    console.log(text);
    
    const res = await callCAIOReviewService(docid, text);
    if (res != null) {
      const reviewed_text = res.data;
      console.log(reviewed_text);

      contentParagraph.insertText(reviewed_text, Word.InsertLocation.replace);

      await context.sync();
    }
    
  } else {
    console.log("Not reviewing :" + text);
  }
  return true;
}

export async function run_rfp_extract() {
  var docid = sessionStorage.getItem("docid");
  const res = await callExtractRFPDataService(docid);
  if (res != null) {
    const data = JSON.parse(res);
    await insertDataTable(data);
    showStatusMessage("Added RFP metadata to the RFP document");
  } else {
    showStatusMessage("Error extracting metadata from the RFP document");
  }
}

export async function add_docid_header(docid) {
  return Word.run(async (context) => {
    const sections = context.document.sections;
    context.load(sections, "body/style");

    await context.sync();

    const myHeader = sections.items[0].getHeader("primary");
    myHeader.insertText(docid, Word.InsertLocation.end);

    await context.sync();
  });
}

export async function generate_tasklist() {
  return Word.run(async (context) => {
    showSpinner();
    toggleDiv();
    const docid = sessionStorage.getItem("docid");

    const res = await callCAIOTaskService(docid);
    if (res != null) {
      console.log(res.data);
      sessionStorage.setItem("tasks", JSON.stringify(res.data));
      renderTasks(res.data);

      hideSpinner();
      document.getElementById("landing-page").classList.add("hidden");
      document.getElementById("app").classList.remove("hidden");
      document.getElementById("app").style.display = "block";
    }
  });
}

export async function runTask(taskIndex) {
  console.log(taskIndex);
  const taskList = JSON.parse(sessionStorage.getItem("tasklist"));
  console.log(taskList);
  const taskname = taskList[taskIndex].title;
  console.log("Task " + taskname + " clicked");

  var docid = sessionStorage.getItem("docid");
  console.log("Starting the task");
  if (taskIndex == 0) {
    console.log("Running metadata extraction task");
    await run_rfp_extract();
  } else if (taskIndex == 1) {
    console.log("Running RFP summary task");
    await run_summary();
  } else if (taskIndex == 2) {
    console.log("Running outline generation task");
    await run_outline();
  } else if (taskIndex == 3) {
    console.log("Running content generation task");
    const tasks = JSON.parse(sessionStorage.getItem("tasks"));
    const options = tasks[taskIndex].options;
    var selected = false;
    if (options != null) {
      for (let i = 0; i < options.length; i++) {
        var option = options[i];
        if (option.selected) {
          selected = true;
          console.log("Generating content for " + option.label);
          await run_content(option.label);
        }
      }
      if (!selected) {
        console.log("No sections selected for generating content");
        showStatusMessage("No section selected for generating content", "error");
      }
    } else {
      console.log("No section list available for generating the content");
      showStatusMessage("No section list available for generating the content", "error");
    }
  } else if (taskIndex == 4) {
    console.log("Running review task");
    const tasks = JSON.parse(sessionStorage.getItem("tasks"));
    const options = tasks[taskIndex].options;
    selected = false;
    if (options != null) {
      for (let i = 0; i < options.length; i++) {
        option = options[i];
        if (option.selected) {
          selected = true;
          console.log("Reviewing content for " + option.label);
          await run_review(option.label);
        }
      }
      if (!selected) {
        console.log("No sections selected for reviewing content");
        showStatusMessage("No section selected for reviewing content", "error");
      }
    } else {
      console.log("No section list available for reviewing the content");
      showStatusMessage("No section list available for reviewing the content", "error");
    }
  } else {
    const tasks = JSON.parse(sessionStorage.getItem("tasks"));
    console.log(tasks);
  }
  await callCAIOTaskUpdateService(docid, { title: taskname, completed: 1 });
  console.log(sessionStorage.getItem("tasks"));
  const tasks = JSON.parse(sessionStorage.getItem("tasks"));
  console.log(tasks);
  const task = tasks[taskIndex];
  task.status = "completed";
  task.completed = true;
  tasks.splice(taskIndex, 1, task);
  // hideSpinner();
  sessionStorage.setItem("tasks", JSON.stringify(tasks));
  renderTasks(tasks);
  updateProgress(tasks);
  updateProgressLabel("");
  checkAllTasksCompleted();
}

export async function run_chat(message) {
  const docid = sessionStorage.getItem("docid");
  const response = await callCAIOChatService(docid, message);
  console.log(response);
  addMessage(response, "received");
}

export async function run_files_upload(files) {
  // showSpinner();
  const docid = sessionStorage.getItem("docid");
  console.log(docid);
  const response = await callUploadFilesService(docid, files);
  console.log(response);
  // hideSpinner();
  return response;
}

export async function summarize_doc(filepath) {
  // showSpinner();
  const docid = sessionStorage.getItem("docid");
  console.log(docid);
  console.log(filepath);
  const response = await callDocumentSummaryService(docid, filepath);
  console.log(response);
  // hideSpinner();
  return response;
}

export async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.log(error);
  }
}

// Add options to task
function add_task_options(taskId, options) {
  const tasks = JSON.parse(sessionStorage.getItem("tasks"));
  const task = tasks[taskId - 1];
  console.log("Adding options for task " + task.title);
  console.log(options);
  if (taskId == 4) {
    task.options = options;
    tasks[4].options = [];
  }

  if (taskId == 5) {
    var existing_options = task.options;
    var counter = 0;
    if (existing_options != null && existing_options.length > 0) {
      counter = existing_options.length;
    }

    if (existing_options == null) {
      task.options = options;
    } else {
      const existing_labels = existing_options.map((item) => item["label"]);
      for (let i = 0; i < options.length; i++) {
        if (!existing_labels.includes(options[i].label)) {
          options[i].id = counter + i;
          existing_options.push(options[i]);
        }
      }

      task.options = existing_options;
    }
  }

  tasks.splice(taskId - 1, 1, task);
  console.log(tasks);
  sessionStorage.setItem("tasks", JSON.stringify(tasks));
}

// Render all tasks
function renderTasks(tasks) {
  const taskList = document.getElementById("task-list");
  taskList.innerHTML = "";

  var counter = 0;
  tasks.forEach((task) => {
    counter = counter + 1;
    const taskElement = createTaskElement(counter, task);
    taskList.appendChild(taskElement);
  });
  updateProgress(tasks);
  checkAllTasksCompleted();
}

// Show spinner overlay
function showSpinner() {
  const spinnerOverlay = document.getElementById('spinner-overlay');
  const app = document.getElementById('app');
  spinnerOverlay.classList.add('active');
  app.classList.add('panel-disabled');
}

// Hide spinner overlay
function hideSpinner() {
  const spinnerOverlay = document.getElementById('spinner-overlay');
  const app = document.getElementById('app');
  spinnerOverlay.classList.remove('active');
  app.classList.remove('panel-disabled');
}

// Handle running a task
export const handleRunTask = (taskId) => {
  const tasks = JSON.parse(sessionStorage.getItem("tasks"));
  const task = tasks[taskId - 1];
  if (task && !task.completed) {
    task.status = "running";
    // showSpinner();
    tasks.splice(taskId - 1, 1, task);
    sessionStorage.setItem("tasks", JSON.stringify(tasks));
    // Simulate task running
    renderTasks(tasks);
    updateProgressLabel(task.title);
    runTask(taskId - 1);
  }
};

// Handle completing a task
export const handleTaskComplete = (taskId) => {
  const tasks = JSON.parse(sessionStorage.getItem("tasks"));
  const task = tasks[taskId - 1];
  if (task) {
    if (task.completed) {
      task.completed = 0;
    } else {
      task.completed = 1;
    }
    // task.completed = !task.completed;
    if (task.completed) {
      task.status = "completed";
    } else {
      task.status = "pending";
    }

    tasks.splice(taskId - 1, 1, task);
    sessionStorage.setItem("tasks", JSON.stringify(tasks));
    renderTasks(tasks);
    updateProgress(tasks);
    checkAllTasksCompleted();
  }
};

// Create a single task element
function createTaskElement(taskid, task) {
  console.log(task);
  const taskItem = document.createElement("div");
  if (task.completed) {
    task.status = "completed";
  } else {
    if (task.status != "running") {
      task.status = "pending";
    }
  }
  taskItem.classList.add("task-item");
  if (task.status == "running") {
    taskItem.classList.add("running");
  }

  const inputNode = document.createElement("input");
  inputNode.type = "checkbox";
  inputNode.classList.add("task-checkbox");
  inputNode.classList.add("ms-Checkbox");
  if (task.completed) {
    inputNode.checked = true;
  }
  if (task.status == "running") {
    inputNode.disabled = true;
  }
  inputNode.addEventListener("click", () => {
    handleTaskComplete(taskid);
  });
  taskItem.appendChild(inputNode);

  const contentNode = document.createElement("div");
  contentNode.classList.add("task-content");
  if (task.completed) {
    contentNode.classList.add("completed");
  }

  contentNode.innerHTML = `
      <h3 class="task-title ms-font-m">${task.title}</h3>
      <div class="task-status">${task.status}</div>
  `;

  taskItem.appendChild(contentNode);

  const actionsNode = document.createElement("div");
  actionsNode.classList.add("task-actions");

  // Add option button for task having options
  console.log(task.options);
  if (task.options != null && task.options.length > 0) {
    const optionsActionNode = document.createElement("button");
    optionsActionNode.classList.add("ms-Button");
    optionsActionNode.classList.add("ms-Button--icon");
    optionsActionNode.title = "Options";
    if (task.status == "running") {
      optionsActionNode.disabled = true;
    }
    optionsActionNode.addEventListener("click", () => {
      handleOptionsClick(taskid);
    });
    optionsActionNode.innerHTML = `
      <i class="ms-Icon ms-Icon--Settings"></i>
    `;
    actionsNode.appendChild(optionsActionNode);
  }

  // Add run button for task not completed
  if (!task.completed) {
    const runActionNode = document.createElement("button");
    runActionNode.classList.add("ms-Button");
    runActionNode.classList.add("ms-Button--icon");
    if (task.status == "running") {
      runActionNode.classList.add("running");
    }
    if (task.status == "running") {
      runActionNode.title = "Running";
    } else {
      runActionNode.title = "Run";
    }
    if (task.status == "running") {
      runActionNode.disabled = true;
    }
    runActionNode.addEventListener("click", () => {
      handleRunTask(taskid);
    });
    if (task.status == "running") {
      runActionNode.innerHTML = `
        <div class="spinner-icon"></div>
      `;
    } else {
      runActionNode.innerHTML = `
        <i class="ms-Icon ms-Icon--Play"></i>
      `;
    }

    actionsNode.appendChild(runActionNode);
  }

  taskItem.appendChild(actionsNode);

  return taskItem;
}

// Update progress bar
function updateProgress(tasks) {
  const totalTasks = tasks.length;
  const completedTasks = tasks.filter((task) => task.completed).length;
  const progressPercentage = (completedTasks / totalTasks) * 100;

  const progressBar = document.getElementById('progress-bar');
  progressBar.style.width = `${progressPercentage}%`;
}

function updateProgressLabel(taskTitle) {
  const progressLabel = document.getElementById('progress-label');
  progressLabel.textContent = taskTitle ? `Running: ${taskTitle}` : "";
}

// Initialize chat functionality
function initializeChat() {
  const chatButton = document.getElementById('chat-button');
  const chatPanel = document.getElementById('chat-panel');
  const closeChat = document.getElementById('close-chat');
  const messageInput = document.getElementById('message-input');
  const sendMessage = document.getElementById('send-message');
  const chatMessages = document.getElementById('chat-messages');

  chatButton.addEventListener('click', () => {
    console.log("open chat window");
    chatPanel.classList.add('open');
  });

  closeChat.addEventListener('click', () => {
    chatPanel.classList.remove('open');
  });

  sendMessage.addEventListener('click', () => {
    const message = messageInput.value.trim();
    if (message) {
      addMessage(message, "sent");
      messageInput.value = "";
      run_chat(message);
    }
  });

  messageInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      sendMessage.click();
    }
  });
}

// Add a message to the chat panel
function addMessage(text, type) {
  const chatMessages = document.getElementById('chat-messages');
  const messageElement = document.createElement('div');
  messageElement.className = `chat-message ${type}`;
  messageElement.textContent = text;
  chatMessages.appendChild(messageElement);
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

// Show options panel
function showOptionsPanel(taskId, task) {
  const optionsPanel = document.createElement("div");
  optionsPanel.className = "options-panel";

  const optionsPanelContainer = document.createElement("div");
  optionsPanelContainer.className = "options-panel-content";

  const optionsPanelHeader = document.createElement("div");
  optionsPanelHeader.className = "options-header";

  const optionsPanelHeaderTitle = document.createElement("h2");
  optionsPanelHeaderTitle.className = "ms-font-l";
  optionsPanelHeaderTitle.innerHTML = "Configure Task: " + task.title;

  const optionPanelHeaderButton = document.createElement("button");
  optionPanelHeaderButton.classList.add("ms-Button");
  optionPanelHeaderButton.classList.add("ms-Button--icon");
  optionPanelHeaderButton.addEventListener("click", () => {
    closeOptionsPanel();
  });
  optionPanelHeaderButton.innerHTML = `
    <i class="ms-Icon ms-Icon--ChromeClose"></i>
  `;

  optionsPanelHeader.appendChild(optionsPanelHeaderTitle);
  optionsPanelHeader.appendChild(optionPanelHeaderButton);

  optionsPanelContainer.appendChild(optionsPanelHeader);

  const optionList = document.createElement("div");
  optionList.className = "options-list";

  for (let i = 0; i < task.options.length; i++) {
    const option = task.options[i];

    const optionItem = document.createElement("div");
    optionItem.className = "option-item";

    const optionLabel = document.createElement("label");
    optionLabel.className = "ms-Checkbox";
    // optionLabel.classList.add("option-item-checkbox");
    // optionLabel.classList.add("ms-Checkbox");

    const inputElement = document.createElement("input");
    inputElement.type = "checkbox";
    inputElement.classList.add("option-item-checkbox");
    inputElement.classList.add("ms-Checkbox-input");
    // inputElement.className = "ms-Checkbox-input";

    if (option.selected) {
      inputElement.checked = true;
    }

    inputElement.addEventListener("change", () => {
      handleOptionChange(taskId, option.id, inputElement.checked);
    });

    optionLabel.appendChild(inputElement);

    const spanElement = document.createElement("span");
    spanElement.className = "ms-Checkbox-label";
    spanElement.innerHTML = option.label;

    optionLabel.appendChild(spanElement);

    optionItem.appendChild(optionLabel);

    optionList.appendChild(optionItem);
  }

  optionsPanelContainer.appendChild(optionList);

  const optionsPanelFooter = document.createElement("div");
  optionsPanelFooter.className = "options-footer";
  const optionPanelFooterButton = document.createElement("button");
  optionPanelFooterButton.classList.add("ms-Button");
  optionPanelFooterButton.classList.add("ms-Button--primary");
  optionPanelFooterButton.addEventListener("click", () => {
    console.log("Options saved");
    closeOptionsPanel();
    // handleRunTask(task.id);
  });
  optionPanelFooterButton.innerHTML = `
    <span class="ms-Button-label">Close</span>
  `;

  optionsPanelFooter.appendChild(optionPanelFooterButton);

  optionsPanelContainer.appendChild(optionsPanelFooter);

  optionsPanel.appendChild(optionsPanelContainer);

  document.body.appendChild(optionsPanel);

  // Animate panel in
  setTimeout(() => optionsPanel.classList.add('active'), 10);
}

// Close options panel
function closeOptionsPanel() {
  const panel = document.querySelector('.options-panel');
  if (panel) {
    panel.classList.remove('active');
    setTimeout(() => panel.remove(), 300);
  }
}

// Handle option change
function handleOptionChange(taskId, optionId, checked) {
  const tasks = JSON.parse(sessionStorage.getItem("tasks"));
  const task = tasks[taskId - 1];
  // const task = tasks.find(t => t.id === taskId);
  var options = task.options;
  if (task && task.options) {
    const option = task.options.find((o) => o.id === optionId);
    if (option) {
      option.selected = checked;
      options.splice(optionId - 1, 1, option);
      console.log(options);
      task.options = options;
      tasks.splice(taskId - 1, 1, task);
      sessionStorage.setItem("tasks", JSON.stringify(tasks));
    }
  }
}

function handleOptionsClick(taskId) {
  const tasks = JSON.parse(sessionStorage.getItem("tasks"));
  const task = tasks[taskId - 1];
  // const task = tasks.find(t => t.id === taskId);
  if (task && task.options) {
    showOptionsPanel(taskId, task);
  }
}

// Document Repo

// Initialize documents repository functionality
function initializeDocumentsRepo() {
  const documentsRepoButton = document.getElementById("documents-repo-button");
  const documentsPanel = document.getElementById("documents-panel");
  const closeDocuments = document.getElementById("close-documents");
  const documentUpload = document.getElementById("file-upload");

  documentsRepoButton.addEventListener("click", () => {
    documentsPanel.classList.add("open");
    renderDocuments();
  });

  closeDocuments.addEventListener("click", () => {
    documentsPanel.classList.remove("open");
  });

  documentUpload.addEventListener("change", (e) => {
    var documents = [];
    if (sessionStorage.getItem("documents") != null) {
      documents = JSON.parse(sessionStorage.getItem("documents"));
    }
    const files = Array.from(e.target.files);
    files.forEach((file) => {
      // Add file to documents array
      const newDocument = {
        id: documents.length + 1,
        name: file.name,
        type: file.type,
        size: formatFileSize(file.size),
        date: new Date().toLocaleDateString(),
        file: file, // Store the actual file object
      };
      run_files_upload(files);
      documents.push(newDocument);
    });

    sessionStorage.setItem("documents", JSON.stringify(documents));

    renderDocuments();
    // renderTasks();
    // updateProgress();

    // Reset file input
    documentUpload.value = "";
  });
}

// Format file size to human-readable format
function formatFileSize(bytes) {
  if (bytes === 0) return "0 Bytes";

  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));

  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

// Render documents list
function renderDocuments() {
  const documentsList = document.getElementById("documents-list");
  var documents = [];
  if (sessionStorage.getItem("documents") != null) {
    documents = JSON.parse(sessionStorage.getItem("documents"));
  }

  if (documents.length === 0) {
    documentsList.innerHTML = `
      <div class="empty-state">
        <i class="ms-Icon ms-Icon--DocumentSet"></i>
        <p>No documents yet. Upload some files to get started.</p>
      </div>
    `;
    return;
  }

  documentsList.innerHTML = "";

  documents.forEach((doc) => {
    const docElement = createDocumentElement(doc);
    documentsList.appendChild(docElement);
  });
}

// Create a single document element
function createDocumentElement(doc) {
  const docItem = document.createElement('div');
  docItem.className = 'document-item';

  // Determine icon based on file type
  let iconClass = 'ms-Icon--Document';
  const doc_type = doc.doc_path.split(".")[1];
  console.log(doc_type);
  if (doc_type == 'pdf') {
    iconClass = 'ms-Icon--PDF';
  } else if (doc_type == 'docx' || doc_type == 'doc') {
    iconClass = 'ms-Icon--WordDocument';
  }

  docItem.innerHTML = `
    <div class="document-icon">
      <i class="ms-Icon ${iconClass}"></i>
    </div>
    <div class="document-info">
      <h3 class="document-name">${doc.doc_path}</h3>
    </div>
  `;

  return docItem;
}

// Handle viewing a document
function handleViewDocument (docId) {
  const documents = JSON.parse(sessionStorage.getItem("documents"));
  const doc = documents.find((d) => d.id === docId);
  if (doc) {
    // In a real application, this would open the document
    alert(`Viewing document: ${doc.name}`);
  }
}

// Handle deleting a document
function handleDeleteDocument (docId) {
  const documents = JSON.parse(sessionStorage.getItem("documents"));
  const docIndex = documents.findIndex((d) => d.id === docId);
  if (docIndex !== -1) {
    documents.splice(docIndex, 1);
    renderDocuments();
  }
}

// Documents Dialog Box

// Sample documents data

// Required document state
let hasRfp = false;
let hasSample = false;

// Initialize upload dialog
function initializeUploadDialog() {
  let files = [];
  const uploadDiv = document.getElementById("upload-dialog");
  const uploadButtons = uploadDiv.querySelectorAll(".upload-button");
  const continueButton = document.getElementById("continue-button");

  uploadButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const inputId = button.getAttribute("data-for");
      document.getElementById(inputId).click();
    });
  });

  // RFP document upload
  document.getElementById("rfp-upload").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) {
      hasRfp = true;
      document.getElementById("rfp-filename").textContent = file.name;
      files.push({
        type: "rfp",
        file: file,
      });
      checkRequiredDocuments();
    }
  });

  // Sample document upload
  document.getElementById("sample-upload").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) {
      hasSample = true;
      document.getElementById("sample-filename").textContent = file.name;
      files.push({
        type: "sample",
        file: file,
      });
      checkRequiredDocuments();
    }
  });

  // Reference documents upload
  document.getElementById("reference-upload").addEventListener("change", (e) => {
    const ref_files = Array.from(e.target.files);
    const fileList = document.getElementById("reference-files");

    ref_files.forEach((file) => {
      files.push({
        type: "ref",
        file: file,
      });
      const fileItem = document.createElement("div");
      fileItem.className = "file-item";
      fileItem.innerHTML = `
        <i class="ms-Icon ms-Icon--PDF"></i>
        <span>${file.name}</span>
      `;
      fileList.appendChild(fileItem);
    });

    checkRequiredDocuments();
  });

  // Continue button click
  continueButton.addEventListener("click", () => {
    document.getElementById("upload-dialog").classList.remove("open");
    document.getElementById("landing-page").classList.add("hidden");
    document.getElementById("app").classList.remove("hidden");

    console.log("Uploading and indexing documents.");
    start_rfp_draft(files);
  });
}

async function start_rfp_draft(files) {
  console.log("Uploading and indexing documents.");
  const loadingScreen = document.getElementById("loading-screen");
  const loadingStatus = document.getElementById("loading-status");
  loadingScreen.classList.remove("hidden");
  loadingStatus.textContent = "Uploading and indexing documents...";
  const response = await run_files_upload(files);
  var templateFile = null;
  for (let k=0; k < files.length; k++) {
    var file = files[k];
    if (file["type"] == "sample") {
      templateFile = file["file"];
    }
  }

  if (response != null) {
    console.log(response);
    const documents = response.data;
    if (documents != null) {
      sessionStorage.setItem("documents", JSON.stringify(documents));
    }
    console.log("Summarizing RFP and Sample documents...");
    // loadingStatus.textContent = "Summarizing documents...";
    for (let i = 0; i < documents.length; i++) {
      const document = documents[i];
      if (document["doc_type"] != "ref") {
        console.log("Summarizing documents... " + document["doc_path"]);
        // loadingStatus.textContent = "Summarizing documents...\r\n" + document["doc_path"];
        loadingStatus.innerHTML = "Summarizing documents...<br />" + document["doc_path"];
        const res = await summarize_doc(document["doc_type"] + "/" + document["doc_path"]);
        console.log(res);
      }
    }
    console.log("Applying template");
    loadingStatus.textContent = "Applying template...";
    await apply_template(templateFile);

    console.log("Generating task list.");
    loadingStatus.textContent = "Generating tasklist...";
    const docid = sessionStorage.getItem("docid");
    await add_docid_header(docid);
    await generate_tasklist();

    loadingScreen.classList.add("hidden");
  } else {
    console.log("Error adding files");
    loadingScreen.classList.add("hidden");
    document.getElementById("landing-page").classList.remove("hidden");
    // style.display = "block";
  }
}

async function complete_rfp_draft(docid) {
  document.getElementById("complete-drafting").disabled = true;
  showSpinner();
  var res = await callCompleteDraftingService(docid);
  console.log(res);
  if (res != null) {
    console.log("Cleaned server resources for the document " + docid);
    showStatusMessage("Drafting completed successfully");
  } else {
    console.log("Error cleaning server resources for the document " + docid);
    showStatusMessage("Error cleaning server resources for the document", "error");
  }
  res = await callDraftStstusService(docid);
  hideSpinner();
  const draft = res["data"];
  renderCompletionPage(draft);
}

// Check if required documents are uploaded
function checkRequiredDocuments() {
  const continueButton = document.getElementById("continue-button");
  continueButton.disabled = !(hasRfp && hasSample);
}

// Check if all tasks are completed
function checkAllTasksCompleted() {
  const tasks = JSON.parse(sessionStorage.getItem("tasks"));
  const allCompleted = tasks.every((task) => task.completed);
  document.getElementById("complete-drafting").disabled = !allCompleted;
}

// Show status message
function showStatusMessage(message, type = "success") {
  const statusPanel = document.getElementById("status-panel");
  const statusMessage = document.createElement("div");
  statusMessage.className = `status-message ${type}`;

  const icon = type === "success" ? "CheckMark" : "ErrorBadge";

  statusMessage.innerHTML = `
    <i class="ms-Icon ms-Icon--${icon}"></i>
    <span>${message}</span>
  `;

  statusPanel.appendChild(statusMessage);

  // Remove the message after 3 seconds
  setTimeout(() => {
    statusMessage.classList.add("fade-out");
    setTimeout(() => {
      statusMessage.remove();
    }, 300);
  }, 3000);
}

// Import template

async function apply_template(templateFile) {
  try {
    // const docid = sessionStorage.getItem("docid");
    // const template = await callCAIODocAsBase64Service(docid);
    if(templateFile != null) {
      await loadTemplateFile(templateFile);
    // }
    // if (template != null) {
    //   await importTemplate(template);
    } else {
      console.log("No template available");
    }
    

  } catch(error) {
    console.log(error);
    showStatusMessage("Error applying template");
  }
}

// Gets the contents of the selected file.
async function loadTemplateFile(templateFile) {
  try {
    console.log("Loading template file");

    // const myTemplate = document.getElementById("template-file");
    const reader = new FileReader();
    var template = null;
    const text = await new Promise((resolve) => {
      reader.onload = (event) => {
        resolve(reader.result.toString());
      };
      // Remove the metadata before the Base64-encoded string.
      // const startIndex = reader.result.toString().indexOf("base64,");
      // template = reader.result.toString().substring(startIndex + 7);
  
      // // Show the Update section.
      // $("#imported-section").show();
      // console.log("Document template applied");
      // Read the file as a data URL so we can parse the Base64-encoded string.
      reader.readAsDataURL(templateFile);
    });
    
    const startIndex = text.indexOf("base64,");
    template = text.substring(startIndex + 7);

    // Import the template into the document.
    await importTemplate(template);

  } catch(error) {
    console.log("Error loading template");
    console.log(error);
  }
 
}

// Imports the template into this document.
async function importTemplate(template) {
  console.log("Importing template")
  await Word.run(async (context) => {
    // Use the Base64-encoded string representation of the selected .docx file.
    context.document.insertFileFromBase64(template, "Replace", {
      importTheme: true,
      importStyles: true,
      importParagraphSpacing: true,
      importPageColor: true,
      importDifferentOddEvenPages: true
    });
    await context.sync();
    console.log("Document template applied");
  });
}