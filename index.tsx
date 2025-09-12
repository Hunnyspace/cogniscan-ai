
import { GoogleGenAI } from "@google/genai";
import { marked } from "marked";

// Fix: Declare pdfjsLib to avoid 'Cannot find name' error.
declare const pdfjsLib: any;

// Fix: Add ambient declaration for the html-docx-js-typescript library.
declare const htmlDocx: any;

// Set the worker source for PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

// Get DOM elements
// Fix: Cast elements to their specific types to access properties like 'value' and 'disabled'.
const pdfFileEl = document.getElementById('pdfFile') as HTMLInputElement;
const pageNumbersEl = document.getElementById('pageNumbers') as HTMLInputElement;
const extractTextBtn = document.getElementById('extractTextBtn') as HTMLButtonElement;
const smartFormatBtn = document.getElementById('smartFormatBtn') as HTMLButtonElement;
const clearBtn = document.getElementById('clearBtn') as HTMLButtonElement;
const extractedContentEl = document.getElementById('extractedContent') as HTMLDivElement;
const formattedContentEl = document.getElementById('formattedContent') as HTMLDivElement;

const progressContainerEl = document.getElementById('progressContainer') as HTMLDivElement;
const progressBarEl = document.getElementById('progressBar') as HTMLDivElement;
const processTimeMessageEl = document.getElementById('processTimeMessage') as HTMLParagraphElement;
const errorMessageEl = document.getElementById('errorMessage') as HTMLParagraphElement;
const successMessageEl = document.getElementById('successMessage') as HTMLParagraphElement;
const apiStatusMessageEl = document.getElementById('apiStatusMessage') as HTMLParagraphElement;

const copyFormattedTextBtn = document.getElementById('copyFormattedTextBtn') as HTMLButtonElement;
const reformatSelectionBtn = document.getElementById('reformatSelectionBtn') as HTMLButtonElement;
const summarizeBtn = document.getElementById('summarizeBtn') as HTMLButtonElement;
const downloadBtn = document.getElementById('downloadBtn') as HTMLButtonElement;
const downloadOptions = document.getElementById('downloadOptions') as HTMLDivElement;
const downloadMdBtn = document.getElementById('downloadMdBtn') as HTMLButtonElement;
const downloadHtmlBtn = document.getElementById('downloadHtmlBtn') as HTMLButtonElement;
const downloadDocxBtn = document.getElementById('downloadDocxBtn') as HTMLButtonElement;


const summaryModalEl = document.getElementById('summaryModal') as HTMLDivElement;
const summaryContentEl = document.getElementById('summaryContent') as HTMLDivElement;
const closeSummaryModalBtn = document.getElementById('closeSummaryModalBtn') as HTMLButtonElement;
const copySummaryBtn = document.getElementById('copySummaryBtn') as HTMLButtonElement;

const extractedPageCounterEl = document.getElementById('extractedPageCounter') as HTMLSpanElement;
const prevExtractedPageBtn = document.getElementById('prevExtractedPageBtn') as HTMLButtonElement;
const nextExtractedPageBtn = document.getElementById('nextExtractedPageBtn') as HTMLButtonElement;
const formattedPageCounterEl = document.getElementById('formattedPageCounter') as HTMLSpanElement;
const prevFormattedPageBtn = document.getElementById('prevFormattedPageBtn') as HTMLButtonElement;
const nextFormattedPageBtn = document.getElementById('nextFormattedPageBtn') as HTMLButtonElement;
const findInputExtracted = document.getElementById('findInputExtracted') as HTMLInputElement;
const findBtnExtracted = document.getElementById('findBtnExtracted') as HTMLButtonElement;
const findInputFormatted = document.getElementById('findInputFormatted') as HTMLInputElement;
const findBtnFormatted = document.getElementById('findBtnFormatted') as HTMLButtonElement;

const guideEl = document.getElementById('guide') as HTMLDivElement;
const closeGuideBtn = document.getElementById('closeGuideBtn') as HTMLButtonElement;

const actionButtons = [extractTextBtn, smartFormatBtn, clearBtn, reformatSelectionBtn, copyFormattedTextBtn, downloadBtn, summarizeBtn];

// App state
// Fix: Add types for state variables.
let pdfDocument: any = null;
let originalWords: string[] = [];
let wordSpans: HTMLSpanElement[] = []; // Store the span elements for quick access
let selectionStartIndex: number = -1;
let selectionEndIndex: number = -1;
let extractedPagesContent: Record<number, string> = {};
let formattedPagesContent: Record<number, string> = {};
let extractedCurrentPage: number = 1;
let formattedCurrentPage: number = 1;

// Initialize Google GenAI
let ai;
try {
    // Fix: The API key must be obtained from process.env.API_KEY.
    ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
} catch (error) {
    console.error("Failed to initialize GoogleGenAI:", error);
    showMessage(errorMessageEl, "Could not initialize AI services. Please check your API key configuration.", "error");
}


// --- UI Helper Functions ---

function showMessage(element: HTMLElement, message: string, type: string = 'loading') {
    element.textContent = message;
    element.classList.remove('hidden');
    if (type === 'error') {
        element.className = 'text-center text-red-600 font-medium mb-4';
    } else if (type === 'success') {
        element.className = 'text-center text-green-600 font-medium mb-4';
    } else {
        element.className = 'text-center text-gray-600 font-medium mb-4';
    }
}

function hideMessage(element: HTMLElement) {
    element.classList.add('hidden');
    element.textContent = '';
}

function clearAllMessages() {
    hideMessage(errorMessageEl);
    hideMessage(successMessageEl);
    hideMessage(processTimeMessageEl);
    apiStatusMessageEl.textContent = '';
    progressContainerEl.classList.add('hidden');
}

function toggleButtons(enable: boolean) {
    actionButtons.forEach(btn => btn.disabled = !enable);
}

// --- Core PDF and Text Functions ---

// Fix: Add types to function parameters and return value.
function parsePageNumbers(input: string): number[] {
    const ranges = input.split(',').map(s => s.trim()).filter(Boolean);
    // Fix: Specify Set type.
    const pages = new Set<number>();
    for (const range of ranges) {
        if (range.includes('-')) {
            let [start, end] = range.split('-').map(Number);
            if (!isNaN(start) && !isNaN(end) && start <= end) {
                // Fix: Ensure pdfDocument is not null before accessing its properties.
                if(pdfDocument && end > pdfDocument.numPages) end = pdfDocument.numPages;
                for (let i = start; i <= end; i++) {
                    pages.add(i);
                }
            }
        } else {
            const pageNum = Number(range);
            if (!isNaN(pageNum)) {
                pages.add(pageNum);
            }
        }
    }
    return Array.from(pages).sort((a, b) => a - b);
}

function updateExtractedView() {
    const pageNum = extractedCurrentPage;
    const pageContent = extractedPagesContent[pageNum];
    const pageKeys = Object.keys(extractedPagesContent).map(Number);

    if (!pageContent) {
        extractedContentEl.innerHTML = 'Page not found.';
        extractedPageCounterEl.textContent = '0/0';
        prevExtractedPageBtn.disabled = true;
        nextExtractedPageBtn.disabled = true;
        return;
    }

    extractedContentEl.innerHTML = '';
    originalWords = [];
    wordSpans = [];
    const pageTextDiv = document.createElement('div');
    const wordsOnPage = pageContent.trim().split(/\s+/).filter(Boolean);
    const fragment = document.createDocumentFragment();
    wordsOnPage.forEach(word => {
        const span = document.createElement('span');
        span.textContent = word;
        // Fix: Convert number to string for setAttribute
        span.setAttribute('data-word-index', String(originalWords.length));
        span.classList.add('word');
        fragment.appendChild(span);
        fragment.appendChild(document.createTextNode(' '));
        wordSpans.push(span);
        originalWords.push(word);
    });
    pageTextDiv.appendChild(fragment);
    extractedContentEl.appendChild(pageTextDiv);

    extractedPageCounterEl.textContent = `${pageKeys.indexOf(pageNum) + 1}/${pageKeys.length}`;
    prevExtractedPageBtn.disabled = pageNum <= Math.min(...pageKeys);
    nextExtractedPageBtn.disabled = pageNum >= Math.max(...pageKeys);
}

// Fix: Make function async to handle async marked.parse()
async function updateFormattedView() {
    const pageNum = formattedCurrentPage;
    const pageContent = formattedPagesContent[pageNum];
    const pageKeys = Object.keys(formattedPagesContent).map(Number);
    
    if (Object.keys(formattedPagesContent).length === 0) {
        formattedContentEl.innerHTML = '<p class="text-gray-400">The formatted text will appear here.</p>';
        formattedPageCounterEl.textContent = '';
        prevFormattedPageBtn.disabled = true;
        nextFormattedPageBtn.disabled = true;
        return;
    }

    if (!pageContent) {
        formattedContentEl.innerHTML = 'Formatted page not found.';
        formattedPageCounterEl.textContent = '0/0';
        prevFormattedPageBtn.disabled = true;
        nextFormattedPageBtn.disabled = true;
        return;
    }
    // Fix: await marked.parse as it can return a Promise
    formattedContentEl.innerHTML = await marked.parse(pageContent);
    formattedPageCounterEl.textContent = `${pageKeys.indexOf(pageNum) + 1}/${pageKeys.length}`;
    prevFormattedPageBtn.disabled = pageNum <= Math.min(...pageKeys);
    nextFormattedPageBtn.disabled = pageNum >= Math.max(...pageKeys);
}

async function extractTextFromPDF() {
    clearAllMessages();
    await updateFormattedView();
    extractedContentEl.innerHTML = 'Processing...';
    toggleButtons(false);
    const startTime = performance.now();

    if (!pdfDocument) {
        showMessage(errorMessageEl, 'Please upload a PDF first.', 'error');
        extractedContentEl.innerHTML = 'Select a PDF and click \'Extract Text\' to see the results here.';
        toggleButtons(true);
        return;
    }

    const pageNumbersInput = pageNumbersEl.value.trim() === '' ? `1-${pdfDocument.numPages}` : pageNumbersEl.value;
    const pageNumbers = parsePageNumbers(pageNumbersInput);

    if (pageNumbers.length === 0) {
        showMessage(errorMessageEl, 'Please enter valid page numbers (e.g., 1, 3-5).', 'error');
        extractedContentEl.innerHTML = 'Select a PDF and click \'Extract Text\' to see the results here.';
        toggleButtons(true);
        return;
    }

    try {
        extractedPagesContent = {};
        const totalPages = pdfDocument.numPages;
        for (const pageNum of pageNumbers) {
            if (pageNum > 0 && pageNum <= totalPages) {
                const page = await pdfDocument.getPage(pageNum);
                const textContent = await page.getTextContent();
                const pageText = textContent.items.map((item: any) => item.str).join(' ');
                extractedPagesContent[pageNum] = pageText;
            }
        }
        
        if (Object.keys(extractedPagesContent).length === 0) {
            extractedContentEl.innerHTML = 'No text found on the specified pages.';
            showMessage(errorMessageEl, 'Could not find text on the specified pages.', 'error');
        } else {
            extractedCurrentPage = Math.min(...Object.keys(extractedPagesContent).map(Number));
            updateExtractedView();
            const endTime = performance.now();
            const duration = ((endTime - startTime) / 1000).toFixed(2);
            showMessage(processTimeMessageEl, `Extracted ${pageNumbers.length} page(s) in ${duration} seconds.`, 'success');
        }

    } catch (error) {
        console.error('Error extracting text:', error);
        showMessage(errorMessageEl, 'An error occurred while extracting text. Please try a different file.', 'error');
        extractedContentEl.innerHTML = 'Select a PDF and click \'Extract Text\' to see the results here.';
    } finally {
        toggleButtons(true);
    }
}

function handleRightClickSelection(event: MouseEvent) {
    event.preventDefault();
    clearAllMessages();

    const target = event.target as HTMLElement;
    if (target.classList.contains('word')) {
        const wordIndex = parseInt(target.getAttribute('data-word-index')!);

        if (selectionStartIndex === -1) {
            wordSpans.forEach(span => span.classList.remove('bg-blue-200'));
            selectionStartIndex = wordIndex;
            selectionEndIndex = -1;
            target.classList.add('bg-blue-200');
            showMessage(successMessageEl, 'Selection started. Right-click on the end word to finish.', 'success');
        } else {
            selectionEndIndex = wordIndex;
            const start = Math.min(selectionStartIndex, selectionEndIndex);
            const end = Math.max(selectionStartIndex, selectionEndIndex);

            wordSpans.forEach((span, i) => {
                if (i >= start && i <= end) {
                    span.classList.add('bg-blue-200');
                } else {
                    span.classList.remove('bg-blue-200');
                }
            });
            
            const selectedWords = originalWords.slice(start, end + 1);
            formattedContentEl.innerHTML = `<p>${selectedWords.join(' ')}</p>`;
            showMessage(successMessageEl, 'Specific text extracted successfully!', 'success');
            selectionStartIndex = -1;
            selectionEndIndex = -1;
        }
    }
}

// --- AI Formatting Functions ---

async function smartFormatText() {
    clearAllMessages();
    toggleButtons(false);
    const startTime = performance.now();
    
    const pagesToFormat = Object.keys(extractedPagesContent).map(Number).sort((a,b)=>a-b);
    if (pagesToFormat.length === 0) {
        showMessage(errorMessageEl, 'No text to format. Please extract text first.', 'error');
        toggleButtons(true);
        return;
    }
    if (!ai) {
        showMessage(errorMessageEl, 'AI service is not available.', 'error');
        toggleButtons(true);
        return;
    }

    formattedPagesContent = {};
    formattedCurrentPage = 1;
    progressContainerEl.classList.remove('hidden');
    progressBarEl.style.width = '0%';
    apiStatusMessageEl.textContent = `Starting smart formatting for ${pagesToFormat.length} pages...`;

    for (const [index, pageNum] of pagesToFormat.entries()) {
        const textToFormat = extractedPagesContent[pageNum];
        if (!textToFormat) continue;

        apiStatusMessageEl.textContent = `Formatting page ${pageNum} (${index + 1}/${pagesToFormat.length})...`;
        const userQuery = `Format the following raw text into a well-structured document. Do not remove any content. Use markdown to create appropriate headings, subheadings, bullet points, and numbered lists. Where data appears to be structured (e.g., in columns or rows), format it into a markdown table with headers. Text to format:\n\n${textToFormat}`;

        try {
            const response = await ai.models.generateContent({
                model: 'gemini-2.5-flash',
                contents: userQuery,
                config: {
                    systemInstruction: "You are a professional document formatter. Your goal is to make the provided text clear and readable using Markdown. Identify key sections, lists, and tables and format them accordingly."
                }
            });
            const formattedText = response.text;
            if (formattedText) {
                formattedPagesContent[pageNum] = formattedText;
            } else {
                formattedPagesContent[pageNum] = `*AI failed to provide a format for this page. Original text:*\n\n${textToFormat}`;
                console.warn(`No formatted text returned for page ${pageNum}.`);
            }
        } catch (error) {
            console.error(`API call failed for page ${pageNum}:`, error);
            formattedPagesContent[pageNum] = `*Error formatting this page. Original text:*\n\n${textToFormat}`;
            showMessage(errorMessageEl, `An error occurred while formatting page ${pageNum}.`, 'error');
        }
        progressBarEl.style.width = `${((index + 1) / pagesToFormat.length) * 100}%`;
    }

    apiStatusMessageEl.textContent = '';
    progressContainerEl.classList.add('hidden');
    
    if (Object.keys(formattedPagesContent).length > 0) {
        formattedCurrentPage = Math.min(...Object.keys(formattedPagesContent).map(Number));
        await updateFormattedView(); // It is now async
        const endTime = performance.now();
        const duration = ((endTime - startTime) / 1000).toFixed(2);
        showMessage(processTimeMessageEl, `Smart formatting completed in ${duration} seconds.`, 'success');
    } else {
        showMessage(errorMessageEl, 'Smart formatting failed for all pages.', 'error');
    }
    toggleButtons(true);
}

async function reformatSelectedText() {
    clearAllMessages();
    const selection = window.getSelection();
    if (!selection) return;
    const selectedText = selection.toString().trim();
    
    if (!selectedText) {
        showMessage(errorMessageEl, 'Please select text from the "Formatted Text" box to re-format.', 'error');
        return;
    }
    if (!ai) {
        showMessage(errorMessageEl, 'AI service is not available.', 'error');
        return;
    }

    toggleButtons(false);
    apiStatusMessageEl.textContent = 'Re-formatting selected text...';

    const userQuery = `Reformat the following text. Look for ways to present data in tables or lists to make it more organized and clear. Do not add or remove any content. Text to re-format:\n\n${selectedText}`;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: userQuery,
            config: {
                systemInstruction: "You are a professional document formatter. Your goal is to make the provided text clear and readable using Markdown. Only re-format the provided text; do not add extra commentary."
            }
        });
        const newText = response.text;
        
        if (newText) {
            // Fix: await marked.parse as it can return a Promise
            const newHtmlText = await marked.parse(newText);
            const range = selection.getRangeAt(0);
            range.deleteContents();
            const fragment = range.createContextualFragment(newHtmlText);
            range.insertNode(fragment);
            showMessage(successMessageEl, 'Selected text re-formatted successfully!', 'success');
        } else {
            showMessage(errorMessageEl, 'Failed to re-format. The AI did not return a response.', 'error');
        }
    } catch (error) {
        console.error('Re-format API call failed:', error);
        showMessage(errorMessageEl, 'An error occurred during re-formatting. Please try again.', 'error');
    } finally {
        apiStatusMessageEl.textContent = '';
        toggleButtons(true);
    }
}

async function summarizeContent() {
    clearAllMessages();
    const allFormattedText = Object.keys(formattedPagesContent)
        .map(Number)
        .sort((a, b) => a - b)
        .map(pageNum => formattedPagesContent[pageNum])
        .join('\n\n---\n\n');

    if (!allFormattedText.trim()) {
        showMessage(errorMessageEl, 'No formatted text to summarize. Please format the text first.', 'error');
        return;
    }

    if (!ai) {
        showMessage(errorMessageEl, 'AI service is not available.', 'error');
        return;
    }

    toggleButtons(false);
    apiStatusMessageEl.textContent = 'Generating summary...';

    const userQuery = `Please provide a concise summary of the following document content. Focus on the key points, main conclusions, and any important data presented. The summary should be well-structured, easy to read, and capture the essence of the document.\n\nDOCUMENT CONTENT:\n\n${allFormattedText}`;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: userQuery,
            config: {
                systemInstruction: "You are an expert academic and professional summarizer. Your goal is to create a brief, accurate, and highly readable summary of the provided text."
            }
        });
        const summaryText = response.text;

        if (summaryText) {
            summaryContentEl.innerHTML = await marked.parse(summaryText);
            summaryModalEl.classList.remove('hidden');
        } else {
            showMessage(errorMessageEl, 'The AI could not generate a summary for the provided text.', 'error');
        }
    } catch (error) {
        console.error('Summarization API call failed:', error);
        showMessage(errorMessageEl, 'An error occurred while generating the summary. Please try again.', 'error');
    } finally {
        apiStatusMessageEl.textContent = '';
        toggleButtons(true);
    }
}

// --- Utility and Event Handler Functions ---

async function copyFormattedText() {
    clearAllMessages();
    // Fix: Use Promise.all with async map, and sort numbers correctly.
    const formattedHtmlPromises = Object.keys(formattedPagesContent).map(Number).sort((a, b) => a - b).map(async (pageNum) => {
        return await marked.parse(formattedPagesContent[pageNum]);
    });
    const allFormattedHtmlArray = await Promise.all(formattedHtmlPromises);
    const allFormattedHtml = allFormattedHtmlArray.join('<hr style="page-break-after: always; border: none;">');


    if (!allFormattedHtml.trim()) {
        showMessage(errorMessageEl, 'No formatted text to copy.', 'error');
        return;
    }

    try {
        const blob = new Blob([allFormattedHtml], { type: 'text/html' });
        const clipboardItem = new ClipboardItem({ 'text/html': blob });
        await navigator.clipboard.write([clipboardItem]);

        const originalText = copyFormattedTextBtn.textContent;
        copyFormattedTextBtn.textContent = 'Copied!';
        copyFormattedTextBtn.classList.replace('bg-indigo-600', 'bg-green-500');
        setTimeout(() => {
            copyFormattedTextBtn.textContent = originalText;
            copyFormattedTextBtn.classList.replace('bg-green-500', 'bg-indigo-600');
        }, 2000);
    } catch (err) {
        console.error('Failed to copy formatted text as HTML:', err);
        showMessage(errorMessageEl, 'Could not copy HTML. Modern browser needed.', 'error');
    }
}

function findAndHighlight(container: HTMLElement, query: string) {
    // 1. Clear previous highlights safely
    const highlights = container.querySelectorAll('span.highlight');
    highlights.forEach(span => {
        const parent = span.parentNode;
        if (!parent) return;
        // Replace the span with its own text content
        parent.replaceChild(document.createTextNode(span.textContent || ''), span);
        parent.normalize(); // Merge adjacent text nodes for cleaner DOM
    });

    if (!query || query.trim() === '') return;

    // 2. Highlight new matches
    const regex = new RegExp(query.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'gi');
    const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
    
    const textNodesToProcess: Node[] = [];
    let node;
    while (node = walker.nextNode()) {
        // Simple check to avoid processing nodes inside scripts or styles, if any
        if (node.parentElement?.tagName === 'SCRIPT' || node.parentElement?.tagName === 'STYLE') {
            continue;
        }
        if (node.nodeValue && regex.test(node.nodeValue)) {
            textNodesToProcess.push(node);
        }
    }

    textNodesToProcess.forEach(textNode => {
        if (!textNode.nodeValue || !textNode.parentNode) return;

        const fragment = document.createDocumentFragment();
        let lastIndex = 0;

        textNode.nodeValue.replace(regex, (match, offset: number) => {
            // Add the text before the match
            const beforeText = textNode.nodeValue!.substring(lastIndex, offset);
            if (beforeText) {
                fragment.appendChild(document.createTextNode(beforeText));
            }

            // Add the highlighted match
            const highlightSpan = document.createElement('span');
            highlightSpan.className = 'highlight';
            highlightSpan.textContent = match;
            fragment.appendChild(highlightSpan);

            lastIndex = offset + match.length;
            return match; // required by replace, but we don't use the returned string
        });
        
        // Add any remaining text after the last match
        const afterText = textNode.nodeValue.substring(lastIndex);
        if (afterText) {
            fragment.appendChild(document.createTextNode(afterText));
        }

        // Replace the original text node with the new fragment
        textNode.parentNode.replaceChild(fragment, textNode);
    });
}

function downloadFile(content: string | Blob, fileName: string, mimeType?: string) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

async function resetApp() {
    pdfFileEl.value = '';
    pageNumbersEl.value = '';
    pdfDocument = null;
    originalWords = [];
    wordSpans = [];
    selectionStartIndex = -1;
    selectionEndIndex = -1;
    extractedPagesContent = {};
    formattedPagesContent = {};
    extractedCurrentPage = 1;
    formattedCurrentPage = 1;

    extractedContentEl.innerHTML = 'Select a PDF and click \'Extract Text\' to see the results here.';
    findInputExtracted.value = '';
    findInputFormatted.value = '';
    updateExtractedView();
    await updateFormattedView();
    clearAllMessages();
    showMessage(successMessageEl, 'Application has been reset.', 'success');
    setTimeout(() => hideMessage(successMessageEl), 2000);
}


// --- Event Listeners ---

closeGuideBtn.addEventListener('click', () => {
    guideEl.classList.add('hidden');
});

pdfFileEl.addEventListener('change', (event) => {
    // Fix: cast event.target to HTMLInputElement to access files property.
    const target = event.target as HTMLInputElement;
    if (!target.files) return;
    const file = target.files[0];
    if (!file) return;

    clearAllMessages();
    const reader = new FileReader();
    reader.onload = async (e) => {
        const data = e.target?.result;
        if (!data) return;
        try {
            extractedContentEl.innerHTML = 'Loading PDF...';
            pdfDocument = await pdfjsLib.getDocument({ data }).promise;
            showMessage(successMessageEl, 'PDF loaded successfully. Enter pages and extract text.', 'success');
        } catch (error) {
            console.error('Error loading PDF:', error);
            pdfDocument = null;
            showMessage(errorMessageEl, 'Failed to load PDF. Please ensure it is a valid file.', 'error');
            extractedContentEl.innerHTML = 'Select a PDF and click \'Extract Text\' to see the results here.';
        }
    };
    reader.readAsArrayBuffer(file);
});

extractTextBtn.addEventListener('click', extractTextFromPDF);
smartFormatBtn.addEventListener('click', smartFormatText);
clearBtn.addEventListener('click', resetApp);
// Fix: Pass event to handler to have it typed.
extractedContentEl.addEventListener('contextmenu', (e) => handleRightClickSelection(e));
copyFormattedTextBtn.addEventListener('click', copyFormattedText);
reformatSelectionBtn.addEventListener('click', reformatSelectedText);
summarizeBtn.addEventListener('click', summarizeContent);

downloadBtn.addEventListener('click', () => {
    downloadOptions.classList.toggle('hidden');
});

document.addEventListener('click', (e) => {
    // Fix: Cast e.target to Node for .contains() method.
    if (!downloadBtn.contains(e.target as Node) && !downloadOptions.contains(e.target as Node)) {
        downloadOptions.classList.add('hidden');
    }
});

downloadMdBtn.addEventListener('click', (e) => {
    e.preventDefault();
    // Fix: Correctly sort numeric keys from object.
    const allFormattedMd = Object.keys(formattedPagesContent).map(Number).sort((a, b) => a - b).map(pageNum => {
        return `## Page ${pageNum}\n\n${formattedPagesContent[pageNum]}`;
    }).join('\n\n---\n\n');
    if(allFormattedMd.trim()) {
        downloadFile(allFormattedMd, 'formatted_text.md', 'text/markdown;charset=utf-8');
    } else {
        showMessage(errorMessageEl, 'No formatted text to download.', 'error');
    }
    downloadOptions.classList.add('hidden');
});

// Fix: Make listener async to handle async marked.parse()
downloadHtmlBtn.addEventListener('click', async (e) => {
    e.preventDefault();
    // Fix: Use Promise.all with async map, and sort numbers correctly.
    const formattedHtmlPromises = Object.keys(formattedPagesContent).map(Number).sort((a,b)=>a-b).map(async (pageNum) => {
        return `<h2>Page ${pageNum}</h2>\n${await marked.parse(formattedPagesContent[pageNum])}`;
    });
    const allFormattedHtmlArray = await Promise.all(formattedHtmlPromises);
    const allFormattedHtml = allFormattedHtmlArray.join('<hr style="page-break-after: always; border-top: 1px solid #ccc;">');

    if(allFormattedHtml.trim()) {
        const finalHtml = `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>Formatted PDF Content</title><style>body{font-family: sans-serif; line-height: 1.6;} table{border-collapse: collapse; width: 100%;} th, td{border: 1px solid #ddd; padding: 8px;} th{background-color: #f2f2f2;}</style></head><body>${allFormattedHtml}</body></html>`;
        downloadFile(finalHtml, 'formatted_text.html', 'text/html;charset=utf-8');
    } else {
        showMessage(errorMessageEl, 'No formatted text to download.', 'error');
    }
    downloadOptions.classList.add('hidden');
});

downloadDocxBtn.addEventListener('click', async (e) => {
    e.preventDefault();
    const formattedHtmlPromises = Object.keys(formattedPagesContent).map(Number).sort((a, b) => a - b).map(async (pageNum) => {
        const pageHtml = await marked.parse(formattedPagesContent[pageNum]);
        return `<h2>Page ${pageNum}</h2>\n${pageHtml}`;
    });
    const allFormattedHtmlArray = await Promise.all(formattedHtmlPromises);
    const combinedHtml = `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>Formatted Document</title></head><body>${allFormattedHtmlArray.join('<br page-break-before="always" />')}</body></html>`;

    if (combinedHtml.trim()) {
        try {
            const fileBlob = await htmlDocx.asBlob(combinedHtml);
            downloadFile(fileBlob, 'formatted_document.docx');
        } catch(err) {
            console.error("Error creating docx file:", err);
            showMessage(errorMessageEl, 'Could not create .docx file.', 'error');
        }
    } else {
        showMessage(errorMessageEl, 'No formatted text to download.', 'error');
    }
    downloadOptions.classList.add('hidden');
});


const pageNavListener = async (isNext: boolean, isExtracted: boolean) => {
    const content = isExtracted ? extractedPagesContent : formattedPagesContent;
    let currentPage = isExtracted ? extractedCurrentPage : formattedCurrentPage;
    const updateView = isExtracted ? updateExtractedView : updateFormattedView;

    const pageNums = Object.keys(content).map(Number).sort((a, b) => a - b);
    const currentIndex = pageNums.indexOf(currentPage);

    let newIndex = currentIndex;
    if (isNext && currentIndex < pageNums.length - 1) {
        newIndex++;
    } else if (!isNext && currentIndex > 0) {
        newIndex--;
    }
    
    if (newIndex !== currentIndex) {
        if (isExtracted) {
            extractedCurrentPage = pageNums[newIndex];
        } else {
            formattedCurrentPage = pageNums[newIndex];
        }
        await updateView();
    }
};

prevExtractedPageBtn.addEventListener('click', () => pageNavListener(false, true));
nextExtractedPageBtn.addEventListener('click', () => pageNavListener(true, true));
prevFormattedPageBtn.addEventListener('click', () => pageNavListener(false, false));
nextFormattedPageBtn.addEventListener('click', () => pageNavListener(true, false));
        
findBtnExtracted.addEventListener('click', () => findAndHighlight(extractedContentEl, findInputExtracted.value));
findInputExtracted.addEventListener('keydown', (e) => {
    if(e.key === 'Enter') {
        e.preventDefault();
        findAndHighlight(extractedContentEl, findInputExtracted.value)
    }
});
findBtnFormatted.addEventListener('click', () => findAndHighlight(formattedContentEl, findInputFormatted.value));
findInputFormatted.addEventListener('keydown', (e) => {
    if(e.key === 'Enter') {
        e.preventDefault();
        findAndHighlight(formattedContentEl, findInputFormatted.value)
    }
});

// Summary Modal Listeners
closeSummaryModalBtn.addEventListener('click', () => {
    summaryModalEl.classList.add('hidden');
});

copySummaryBtn.addEventListener('click', async () => {
    const summaryText = summaryContentEl.innerText;
    if (summaryText) {
        try {
            await navigator.clipboard.writeText(summaryText);
            const originalText = copySummaryBtn.textContent;
            copySummaryBtn.textContent = 'Copied!';
            setTimeout(() => {
                copySummaryBtn.textContent = originalText;
            }, 2000);
        } catch (err) {
            console.error('Failed to copy summary:', err);
            alert('Failed to copy summary text.');
        }
    }
});

// Initial Setup
(async () => {
    await updateFormattedView();
})();