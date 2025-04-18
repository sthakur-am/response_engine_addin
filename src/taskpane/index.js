import "./style.css";

import axios from "axios";
import mammoth from "mammoth";
import { jwtDecode } from "jwt-decode";
import { v4 as uuidv4 } from "uuid";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (Office.context.requirements.isSetSupported("WordApi", "1.9")) {
      // Your code that uses Word JavaScript API 1.9 features
      console.log("Word JavaScript API 1.9 is supported.");
    } else {
      console.log("Word JavaScript API 1.9 is not supported.");
    }

    var SpinnerElements = document.querySelectorAll(".ms-Spinner");
    for (var i = 0; i < SpinnerElements.length; i++) {
      new fabric['Spinner'](SpinnerElements[i]);
    }

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = applyTemplate;

    document.getElementById("generateTemplate").onclick = showGenerateTemplateDialog;
    document.getElementById("generateContent").onclick = run_content;

    renderAddIn();
  }
});

var enabled = false;
var dialog = document.getElementById("dialogDiv");
var dialogComponent = new fabric["Dialog"](dialog);

export const showGenerateTemplateDialog = () => {
  console.log("Hi from generateTemplate");
  document.getElementById("selectInputDocsDialog").style.display = "block";
  const dialogAction = new fabric["Button"](document.getElementById("generateActionButton"), generate_tasklist);
  const dialogCancelAction = new fabric["Button"](document.getElementById("cancelActionButton"));

  dialogComponent.open();
};
// Sample tasks data
const tasks = [
  { id: 1, title: 'Review Project Proposal', status: 'pending', completed: false },
  { id: 2, title: 'Update Documentation', status: 'pending', completed: false },
  { id: 3, title: 'Test New Features', status: 'pending', completed: false },
  { id: 4, title: 'Deploy to Production', status: 'pending', completed: false },
];

// Initialize the application
function initializeApp() {
  renderTasks();
  updateProgress();
  initializeChat();
  initializeFileUpload();
}

// Render all tasks
function renderTasks() {
  const taskList = document.getElementById('task-list');
  taskList.innerHTML = '';

  tasks.forEach(task => {
    const taskElement = createTaskElement(task);
    taskList.appendChild(taskElement);
  });
}

// Create a single task element
function createTaskElement(task) {
  const taskItem = document.createElement('div');
  taskItem.className = 'task-item';
  taskItem.innerHTML = `
    <input type="checkbox" class="task-checkbox ms-CheckBox" 
           ${task.completed ? 'checked' : ''} 
           onclick="handleTaskComplete(${task.id})">
    <div class="task-content ${task.completed ? 'completed' : ''}">
      <h3 class="task-title ms-font-m">${task.title}</h3>
      <div class="task-status">${task.status}</div>
    </div>
    <div class="task-actions">
      ${!task.completed ? `
        <button class="ms-Button ms-Button--primary" onclick="handleRunTask(${task.id})">
          <span class="ms-Button-label">Run</span>
        </button>
      ` : ''}
    </div>
  `;
  return taskItem;
}

// Handle running a task
window.handleRunTask = (taskId) => {
  const task = tasks.find(t => t.id === taskId);
  if (task && task.status === 'pending') {
    task.status = 'running';
    
    // Simulate task running
    setTimeout(() => {
      task.status = 'completed';
      task.completed = true;
      renderTasks();
      updateProgress();
    }, 2000);

    renderTasks();
  }
};

// Handle completing a task
window.handleTaskComplete = (taskId) => {
  const task = tasks.find(t => t.id === taskId);
  if (task) {
    task.completed = !task.completed;
    task.status = task.completed ? 'completed' : 'pending';
    renderTasks();
    updateProgress();
  }
};

// Update progress bar
function updateProgress() {
  const totalTasks = tasks.length;
  const completedTasks = tasks.filter(task => task.completed).length;
  const progressPercentage = (completedTasks / totalTasks) * 100;
  
  const progressBar = document.getElementById('progress-bar');
  progressBar.style.width = `${progressPercentage}%`;
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
    chatPanel.classList.add('open');
  });

  closeChat.addEventListener('click', () => {
    chatPanel.classList.remove('open');
  });

  sendMessage.addEventListener('click', () => {
    const message = messageInput.value.trim();
    if (message) {
      addMessage(message, 'sent');
      messageInput.value = '';
      
      // Simulate received message
      setTimeout(() => {
        addMessage('Message received!', 'received');
      }, 1000);
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

// Initialize file upload functionality
function initializeFileUpload() {
  const fileUpload = document.getElementById('file-upload');
  
  fileUpload.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    
    files.forEach(file => {
      // Create a new task for each uploaded file
      const newTask = {
        id: tasks.length + 1,
        title: `Process ${file.name}`,
        status: 'pending',
        completed: false
      };
      
      tasks.push(newTask);
    });
    
    renderTasks();
    updateProgress();
    
    // Reset file input
    fileUpload.value = '';
  });
}

// Initialize the app when the document is loaded
document.addEventListener('DOMContentLoaded', initializeApp);