import { GoogleGenAI, Modality, Type } from "@google/genai";
import { marked } from "marked";
import { asBlob } from 'html-docx-js-typescript';
import { PDFDocument } from 'pdf-lib';
import './index.css';

// Declare third-party libraries loaded via script tags
declare const pdfjsLib: any;
declare const JsBarcode: any;

// Set the worker source for PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

// --- Get DOM Elements ---
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
const downloadInfoModalEl = document.getElementById('downloadInfoModal') as HTMLDivElement;
const closeDownloadInfoModalBtn = document.getElementById('closeDownloadInfoModalBtn') as HTMLButtonElement;
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
const followGateModalEl = document.getElementById('followGateModal') as HTMLDivElement;
const followCheckboxEl = document.getElementById('followCheckbox') as HTMLInputElement;
const continueToAppBtn = document.getElementById('continueToAppBtn') as HTMLButtonElement;

// --- New Feature Elements (v2.0) ---
const bookCoverBtn = document.getElementById('bookCoverBtn') as HTMLButtonElement;
const generateForewordBtn = document.getElementById('generateForewordBtn') as HTMLButtonElement;
const mergePdfsBtn = document.getElementById('mergePdfsBtn') as HTMLButtonElement;

// What's New Modal
const whatsNewModalEl = document.getElementById('whatsNewModal') as HTMLDivElement;
const closeWhatsNewModalBtn = document.getElementById('closeWhatsNewModalBtn') as HTMLButtonElement;
const gotItWhatsNewBtn = document.getElementById('gotItWhatsNewBtn') as HTMLButtonElement;

// Foreword Modal
const forewordModalEl = document.getElementById('forewordModal') as HTMLDivElement;
const closeForewordModalBtn = document.getElementById('closeForewordModalBtn') as HTMLButtonElement;
const forewordContentEl = document.getElementById('forewordContent') as HTMLDivElement;
const copyForewordBtn = document.getElementById('copyForewordBtn') as HTMLButtonElement;

// PDF Merge Modal
const pdfMergeModalEl = document.getElementById('pdfMergeModal') as HTMLDivElement;
const closePdfMergeModalBtn = document.getElementById('closePdfMergeModalBtn') as HTMLButtonElement;
const pdfMergeInput = document.getElementById('pdfMergeInput') as HTMLInputElement;
const executeMergeBtn = document.getElementById('executeMergeBtn') as HTMLButtonElement;

// Book Cover Studio Modal
const bookCoverModalEl = document.getElementById('bookCoverModal') as HTMLDivElement;
const closeBookCoverModalBtn = document.getElementById('closeBookCoverModalBtn') as HTMLButtonElement;
const coverAuthorInput = document.getElementById('coverAuthorInput') as HTMLInputElement;
const coverPriceInput = document.getElementById('coverPriceInput') as HTMLInputElement;
const coverIsbnInput = document.getElementById('coverIsbnInput') as HTMLInputElement;
const generateTitlesBtn = document.getElementById('generateTitlesBtn') as HTMLButtonElement;
const titlesSpinner = document.getElementById('titlesSpinner') as HTMLDivElement;
const titlesOutput = document.getElementById('titlesOutput') as HTMLDivElement;
const coverReferenceImage = document.getElementById('coverReferenceImage') as HTMLInputElement;
const coverPromptInput = document.getElementById('coverPromptInput') as HTMLTextAreaElement;
const generateImageBtn = document.getElementById('generateImageBtn') as HTMLButtonElement;
const imageSpinner = document.getElementById('imageSpinner') as HTMLDivElement;
const coverCanvas = document.getElementById('coverCanvas') as HTMLCanvasElement;
const downloadCoverBtn = document.getElementById('downloadCoverBtn') as HTMLButtonElement;

const actionButtons = [extractTextBtn, smartFormatBtn, clearBtn, reformatSelectionBtn, copyFormattedTextBtn, downloadBtn, summarizeBtn, bookCoverBtn, generateForewordBtn, mergePdfsBtn];

// --- App State ---
const APP_VERSION = '2.0';
let pdfDocument: any = null;
let originalWords: string[] = [];
let wordSpans: HTMLSpanElement[] = []; 
let selectionStartIndex: number = -1;
let selectionEndIndex: number = -1;
let extractedPagesContent: Record<number, string> = {};
let formattedPagesContent: Record<number, string> = {};
let extractedCurrentPage: number = 1;
let formattedCurrentPage: number = 1;
// Book Cover State
let coverState = {
    title: 'Your Book Title',
    author: '',
    price: '',
    isbn: '',
    referenceImage: null as string | null, // Base64
    generatedImage: null as string | null, // Base64
};

// Initialize Google GenAI
let ai;
try {
    ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
} catch (error) {
    console.error("Failed to initialize GoogleGenAI:", error);
    showMessage(errorMessageEl, "Could not initialize AI services. Please check your API key configuration.", "error");
}

// --- UI Helper Functions ---
function showMessage(element: HTMLElement, message: string, type: string = 'loading') {
    element.textContent = message;
    element.classList.remove('hidden');
    element.className = 'text-center font-medium mb-4 ';
    if (type === 'error') element.classList.add('text-red-600');
    else if (type === 'success') element.classList.add('text-green-600');
    else element.classList.add('text-gray-600');
}
function hideMessage(element: HTMLElement) {
    element.classList.add('hidden');
    element.textContent = '';
}
function clearAllMessages() {
    [errorMessageEl, successMessageEl, processTimeMessageEl].forEach(hideMessage);
    apiStatusMessageEl.textContent = '';
    progressContainerEl.classList.add('hidden');
}
function toggleButtons(enable: boolean) {
    actionButtons.forEach(btn => btn.disabled = !enable);
}
const fileToBase64 = (file: File): Promise<string> => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
});

// Generic AI error handler
function handleAiError(error: any, element: HTMLElement) {
    console.error("AI Error:", error);
    const errorMessage = error.toString();
    if (errorMessage.includes('429') || errorMessage.includes('RESOURCE_EXHAUSTED')) {
        showMessage(element, 'API Rate Limit Reached: You have exceeded the free usage quota. Please try again later.', 'error');
    } else {
        showMessage(element, 'An error occurred with the AI service. Please try again.', 'error');
    }
}

// --- Core PDF and Text Functions (Existing) ---
function parsePageNumbers(input: string): number[] {
    const ranges = input.split(',').map(s => s.trim()).filter(Boolean);
    const pages = new Set<number>();
    for (const range of ranges) {
        if (range.includes('-')) {
            let [start, end] = range.split('-').map(Number);
            if (!isNaN(start) && !isNaN(end) && start <= end) {
                if(pdfDocument && end > pdfDocument.numPages) end = pdfDocument.numPages;
                for (let i = start; i <= end; i++) pages.add(i);
            }
        } else {
            const pageNum = Number(range);
            if (!isNaN(pageNum)) pages.add(pageNum);
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
        let totalExtractedText = '';
        const totalPages = pdfDocument.numPages;
        for (const pageNum of pageNumbers) {
            if (pageNum > 0 && pageNum <= totalPages) {
                const page = await pdfDocument.getPage(pageNum);
                const textContent = await page.getTextContent();
                const pageText = textContent.items.map((item: any) => item.str).join(' ');
                extractedPagesContent[pageNum] = pageText;
                totalExtractedText += pageText.trim();
            }
        }
        
        if (Object.keys(extractedPagesContent).length === 0 || totalExtractedText.trim().length === 0) {
            extractedContentEl.innerHTML = '<p class="text-gray-500">No text could be found on the specified pages. This can happen if the PDF contains only images or scanned documents without a text layer.</p>';
            showMessage(errorMessageEl, 'Could not find any text on the specified pages.', 'error');
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
                if (i >= start && i <= end) span.classList.add('bg-blue-200');
                else span.classList.remove('bg-blue-200');
            });
            
            const selectedWords = originalWords.slice(start, end + 1);
            formattedContentEl.innerHTML = `<p>${selectedWords.join(' ')}</p>`;
            showMessage(successMessageEl, 'Specific text extracted successfully!', 'success');
            selectionStartIndex = -1;
            selectionEndIndex = -1;
        }
    }
}
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
        if (!textToFormat || textToFormat.trim().length === 0) {
            formattedPagesContent[pageNum] = "*(This page contained no text to format.)*";
            continue;
        };

        apiStatusMessageEl.textContent = `Formatting page ${pageNum} (${index + 1}/${pagesToFormat.length})...`;
        const userQuery = `Format the following raw text into a well-structured document. Do not remove any content. Use markdown to create appropriate headings, subheadings, bullet points, and numbered lists. Where data appears to be structured (e.g., in columns or rows), format it into a markdown table with headers. Text to format:\n\n${textToFormat}`;

        try {
            const response = await ai.models.generateContent({
                model: 'gemini-2.5-flash',
                contents: userQuery,
                config: { systemInstruction: "You are a professional document formatter. Your goal is to make the provided text clear and readable using Markdown. Identify key sections, lists, and tables and format them accordingly." }
            });
            if (response.text) formattedPagesContent[pageNum] = response.text;
            else {
                formattedPagesContent[pageNum] = `*AI failed to provide a format for this page. Original text:*\n\n${textToFormat}`;
                console.warn(`No formatted text returned for page ${pageNum}.`);
            }
        } catch (error) {
            handleAiError(error, errorMessageEl);
            formattedPagesContent[pageNum] = `*Error formatting this page. Original text:*\n\n${textToFormat}`;
        }
        progressBarEl.style.width = `${((index + 1) / pagesToFormat.length) * 100}%`;
    }

    apiStatusMessageEl.textContent = '';
    progressContainerEl.classList.add('hidden');
    
    if (Object.keys(formattedPagesContent).length > 0) {
        formattedCurrentPage = Math.min(...Object.keys(formattedPagesContent).map(Number));
        await updateFormattedView();
        const endTime = performance.now();
        const duration = ((endTime - startTime) / 1000).toFixed(2);
        showMessage(processTimeMessageEl, `Smart formatting completed in ${duration} seconds.`, 'success');
    } else showMessage(errorMessageEl, 'Smart formatting failed for all pages.', 'error');
    
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
            config: { systemInstruction: "You are a professional document formatter. Your goal is to make the provided text clear and readable using Markdown. Only re-format the provided text; do not add extra commentary." }
        });
        const newText = response.text;
        
        if (newText) {
            const newHtmlText = await marked.parse(newText);
            const range = selection.getRangeAt(0);
            range.deleteContents();
            const fragment = range.createContextualFragment(newHtmlText);
            range.insertNode(fragment);
            showMessage(successMessageEl, 'Selected text re-formatted successfully!', 'success');
        } else showMessage(errorMessageEl, 'Failed to re-format. The AI did not return a response.', 'error');
        
    } catch (error) {
        handleAiError(error, errorMessageEl);
    } finally {
        apiStatusMessageEl.textContent = '';
        toggleButtons(true);
    }
}
async function summarizeContent() {
    clearAllMessages();
    const allFormattedText = Object.values(formattedPagesContent).join('\n\n---\n\n');

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
            config: { systemInstruction: "You are an expert academic and professional summarizer. Your goal is to create a brief, accurate, and highly readable summary of the provided text." }
        });
        const summaryText = response.text;

        if (summaryText) {
            summaryContentEl.innerHTML = await marked.parse(summaryText);
            summaryModalEl.classList.remove('hidden');
        } else showMessage(errorMessageEl, 'The AI could not generate a summary for the provided text.', 'error');
        
    } catch (error) {
        handleAiError(error, errorMessageEl);
    } finally {
        apiStatusMessageEl.textContent = '';
        toggleButtons(true);
    }
}
async function copyFormattedText() {
    clearAllMessages();
    const formattedHtmlPromises = Object.keys(formattedPagesContent).map(Number).sort((a, b) => a - b).map(pageNum => marked.parse(formattedPagesContent[pageNum]));
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
    const highlights = container.querySelectorAll('span.highlight');
    highlights.forEach(span => {
        const parent = span.parentNode;
        if (!parent) return;
        parent.replaceChild(document.createTextNode(span.textContent || ''), span);
        parent.normalize(); 
    });

    if (!query || query.trim() === '') return;

    const regex = new RegExp(query.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'gi');
    const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
    
    const textNodesToProcess: Node[] = [];
    let node;
    while (node = walker.nextNode()) {
        if (node.parentElement?.tagName === 'SCRIPT' || node.parentElement?.tagName === 'STYLE') continue;
        if (node.nodeValue && regex.test(node.nodeValue)) textNodesToProcess.push(node);
    }

    textNodesToProcess.forEach(textNode => {
        if (!textNode.nodeValue || !textNode.parentNode) return;
        const fragment = document.createDocumentFragment();
        let lastIndex = 0;
        textNode.nodeValue.replace(regex, (match, offset: number) => {
            const beforeText = textNode.nodeValue!.substring(lastIndex, offset);
            if (beforeText) fragment.appendChild(document.createTextNode(beforeText));
            const highlightSpan = document.createElement('span');
            highlightSpan.className = 'highlight';
            highlightSpan.textContent = match;
            fragment.appendChild(highlightSpan);
            lastIndex = offset + match.length;
            return match; 
        });
        const afterText = textNode.nodeValue.substring(lastIndex);
        if (afterText) fragment.appendChild(document.createTextNode(afterText));
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

// --- NEW FEATURE FUNCTIONS (v2.0) ---
async function generateForeword() {
    clearAllMessages();
    const allFormattedText = Object.values(formattedPagesContent).join('\n\n');
    if (!allFormattedText.trim()) {
        showMessage(errorMessageEl, 'No formatted text available to generate a foreword.', 'error');
        return;
    }
    if (!ai) {
        showMessage(errorMessageEl, 'AI service is not available.', 'error');
        return;
    }

    toggleButtons(false);
    apiStatusMessageEl.textContent = 'Generating foreword...';

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: `Based on the following document content, write a compelling and professional foreword for a book. The foreword should be about 250-400 words. It should introduce the main topics, highlight the importance of the work, and engage the reader.\n\nDOCUMENT CONTENT:\n${allFormattedText.substring(0, 15000)}`,
            config: { systemInstruction: "You are a professional book editor and author, skilled at writing insightful and engaging forewords." }
        });
        const forewordText = response.text;

        if (forewordText) {
            forewordContentEl.innerHTML = await marked.parse(forewordText);
            forewordModalEl.classList.remove('hidden');
        } else showMessage(errorMessageEl, 'The AI could not generate a foreword for the provided text.', 'error');
        
    } catch (error) {
        handleAiError(error, errorMessageEl);
    } finally {
        apiStatusMessageEl.textContent = '';
        toggleButtons(true);
    }
}
async function mergePdfs() {
    const files = pdfMergeInput.files;
    if (!files || files.length < 2) {
        showMessage(errorMessageEl, 'Please select at least two PDF files to merge.', 'error');
        return;
    }
    apiStatusMessageEl.textContent = `Merging ${files.length} PDFs...`;
    toggleButtons(false);

    try {
        const mergedPdf = await PDFDocument.create();
        for (const file of files) {
            const pdfBytes = await file.arrayBuffer();
            const pdf = await PDFDocument.load(pdfBytes);
            const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
            copiedPages.forEach((page) => mergedPdf.addPage(page));
        }
        const mergedPdfBytes = await mergedPdf.save();
        downloadFile(new Blob([mergedPdfBytes], { type: 'application/pdf' }), 'merged_document.pdf');
        showMessage(successMessageEl, 'PDFs merged successfully!', 'success');
    } catch(err) {
        console.error("Error merging PDFs:", err);
        showMessage(errorMessageEl, 'Could not merge PDFs. Please check the files and try again.', 'error');
    } finally {
        apiStatusMessageEl.textContent = '';
        toggleButtons(true);
        pdfMergeInput.value = '';
        executeMergeBtn.disabled = true;
        pdfMergeModalEl.classList.add('hidden');
    }
}
async function generateTitlesAndCaptions() {
    const allFormattedText = Object.values(formattedPagesContent).join('\n\n');
    if (!allFormattedText.trim() || !ai) {
        showMessage(errorMessageEl, 'Please format text first to generate ideas.', 'error');
        return;
    }
    titlesSpinner.classList.remove('hidden');
    generateTitlesBtn.disabled = true;

    try {
        const response = await ai.models.generateContent({
            model: "gemini-2.5-flash",
            contents: `Based on the following book content, generate 5 catchy, professional book titles and 3 short, engaging back-cover captions. \n\nCONTENT:\n${allFormattedText.substring(0, 10000)}`,
            config: {
                responseMimeType: "application/json",
                responseSchema: {
                    type: Type.OBJECT,
                    properties: {
                        titles: { type: Type.ARRAY, items: { type: Type.STRING } },
                        captions: { type: Type.ARRAY, items: { type: Type.STRING } }
                    }
                },
            },
        });
        
        const suggestions = JSON.parse(response.text);
        titlesOutput.innerHTML = '';
        const titleHeader = document.createElement('h4');
        titleHeader.className = 'font-semibold text-gray-700 mt-2';
        titleHeader.textContent = 'Suggested Titles:';
        titlesOutput.appendChild(titleHeader);
        suggestions.titles.forEach((title: string) => {
            const btn = document.createElement('button');
            btn.className = 'suggestion-btn';
            btn.textContent = title;
            btn.onclick = () => {
                coverState.title = title;
                updateCoverCanvas();
            };
            titlesOutput.appendChild(btn);
        });

        const captionHeader = document.createElement('h4');
        captionHeader.className = 'font-semibold text-gray-700 mt-4';
        captionHeader.textContent = 'Suggested Captions:';
        titlesOutput.appendChild(captionHeader);
        suggestions.captions.forEach((caption: string) => {
            const p = document.createElement('p');
            p.className = 'suggestion-btn';
            p.textContent = `"${caption}"`;
            titlesOutput.appendChild(p);
        });

    } catch (err) {
        handleAiError(err, errorMessageEl);
    } finally {
        titlesSpinner.classList.add('hidden');
        generateTitlesBtn.disabled = false;
    }
}
async function generateCoverImage() {
    if (!coverPromptInput.value.trim() || !ai) {
        showMessage(errorMessageEl, 'Please enter a description for the cover art.', 'error');
        return;
    }
    imageSpinner.classList.remove('hidden');
    generateImageBtn.disabled = true;

    try {
        const parts: any[] = [{ text: coverPromptInput.value }];
        if (coverState.referenceImage) {
            parts.unshift({
                inlineData: {
                    data: coverState.referenceImage.split(',')[1], // remove the data URI prefix
                    mimeType: coverState.referenceImage.match(/data:(.*);/)?.[1] || 'image/jpeg',
                }
            });
        }
        const modelToUse = coverState.referenceImage ? 'gemini-2.5-flash-image-preview' : 'imagen-4.0-generate-001';
        
        let generatedImageData: string | null = null;

        if (modelToUse === 'imagen-4.0-generate-001') {
             const response = await ai.models.generateImages({
                model: 'imagen-4.0-generate-001',
                prompt: coverPromptInput.value,
            });
            generatedImageData = response.generatedImages[0].image.imageBytes;
        } else { // gemini-2.5-flash-image-preview
            const response = await ai.models.generateContent({
                model: 'gemini-2.5-flash-image-preview',
                contents: { parts: parts },
                config: { responseModalities: [Modality.IMAGE, Modality.TEXT] }
            });
            for (const part of response.candidates[0].content.parts) {
                if (part.inlineData) {
                    generatedImageData = part.inlineData.data;
                    break;
                }
            }
        }
        
        if (generatedImageData) {
            coverState.generatedImage = `data:image/png;base64,${generatedImageData}`;
            updateCoverCanvas();
            downloadCoverBtn.disabled = false;
        } else {
             showMessage(errorMessageEl, 'AI failed to generate an image.', 'error');
        }

    } catch(err) {
        handleAiError(err, errorMessageEl);
    } finally {
        imageSpinner.classList.add('hidden');
        generateImageBtn.disabled = false;
    }
}
async function updateCoverCanvas() {
    const ctx = coverCanvas.getContext('2d');
    if (!ctx) return;
    
    // Standard book cover ratio (e.g., 6x9 inches) -> 1200x1800 pixels for high quality
    coverCanvas.width = 1200;
    coverCanvas.height = 1800;

    // Clear canvas
    ctx.fillStyle = '#e5e7eb'; // gray-200
    ctx.fillRect(0, 0, coverCanvas.width, coverCanvas.height);
    
    // Draw generated image
    if (coverState.generatedImage) {
        const img = new Image();
        img.src = coverState.generatedImage;
        await new Promise(resolve => { img.onload = resolve; });
        ctx.drawImage(img, 0, 0, coverCanvas.width, coverCanvas.height);
    } else {
        ctx.fillStyle = 'white';
        ctx.font = '40px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText('Image will appear here', coverCanvas.width / 2, coverCanvas.height / 2);
    }

    // Draw a semi-transparent overlay for text readability
    ctx.fillStyle = 'rgba(0, 0, 0, 0.4)';
    ctx.fillRect(0, 0, coverCanvas.width, coverCanvas.height);

    // Draw Title
    ctx.fillStyle = 'white';
    ctx.textAlign = 'center';
    ctx.font = 'bold 120px serif';
    ctx.fillText(coverState.title, coverCanvas.width / 2, 300, 1000); // Max width 1000px

    // Draw Author
    ctx.font = '60px sans-serif';
    ctx.fillText(coverState.author, coverCanvas.width / 2, 500);

    // Draw Barcode
    if (coverState.price || coverState.isbn) {
        const barcodeCanvas = document.createElement('canvas');
        const barcodeValue = coverState.isbn || coverState.price;
        JsBarcode(barcodeCanvas, barcodeValue, {
            format: coverState.isbn ? "EAN13" : "CODE128",
            background: "#ffffff",
            width: 4,
            height: 150,
            fontSize: 30
        });
        ctx.drawImage(barcodeCanvas, (coverCanvas.width - barcodeCanvas.width) / 2, coverCanvas.height - 300);
    }
}
function assembleAndDownloadCover() {
    updateCoverCanvas().then(() => {
        const dataUrl = coverCanvas.toDataURL('image/png');
        downloadFile(dataUrl, 'book_cover.png');
    });
}


// --- Event Listeners ---
// Existing Listeners
closeGuideBtn.addEventListener('click', () => guideEl.classList.add('hidden'));
pdfFileEl.addEventListener('change', (event) => {
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
extractedContentEl.addEventListener('contextmenu', (e) => handleRightClickSelection(e));
copyFormattedTextBtn.addEventListener('click', copyFormattedText);
reformatSelectionBtn.addEventListener('click', reformatSelectedText);
summarizeBtn.addEventListener('click', summarizeContent);
downloadBtn.addEventListener('click', () => downloadOptions.classList.toggle('hidden'));
document.addEventListener('click', (e) => {
    if (!downloadBtn.contains(e.target as Node) && !downloadOptions.contains(e.target as Node)) {
        downloadOptions.classList.add('hidden');
    }
});
downloadMdBtn.addEventListener('click', (e) => {
    e.preventDefault();
    const allFormattedMd = Object.keys(formattedPagesContent).map(Number).sort((a, b) => a - b).map(pageNum => `## Page ${pageNum}\n\n${formattedPagesContent[pageNum]}`).join('\n\n---\n\n');
    if(allFormattedMd.trim()) downloadFile(allFormattedMd, 'formatted_text.md', 'text/markdown;charset=utf-8');
    else showMessage(errorMessageEl, 'No formatted text to download.', 'error');
    downloadOptions.classList.add('hidden');
});
downloadHtmlBtn.addEventListener('click', async (e) => {
    e.preventDefault();
    // FIX: Use async/await to handle both sync and async return values from marked.parse, preventing a runtime error when calling .then on a string.
    const formattedHtmlPromises = Object.keys(formattedPagesContent).map(Number).sort((a,b)=>a-b).map(async pageNum => `<h2>Page ${pageNum}</h2>\n${await marked.parse(formattedPagesContent[pageNum])}`);
    const allFormattedHtmlArray = await Promise.all(formattedHtmlPromises);
    const allFormattedHtml = allFormattedHtmlArray.join('<hr style="page-break-after: always; border-top: 1px solid #ccc;">');
    if(allFormattedHtml.trim()) {
        const finalHtml = `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>Formatted PDF Content</title><style>body{font-family: sans-serif; line-height: 1.6;} table{border-collapse: collapse; width: 100%;} th, td{border: 1px solid #ddd; padding: 8px;} th{background-color: #f2f2f2;}</style></head><body>${allFormattedHtml}</body></html>`;
        downloadFile(finalHtml, 'formatted_text.html', 'text/html;charset=utf-8');
    } else showMessage(errorMessageEl, 'No formatted text to download.', 'error');
    downloadOptions.classList.add('hidden');
});
downloadDocxBtn.addEventListener('click', async (e) => {
    e.preventDefault();
    downloadInfoModalEl.classList.remove('hidden');
    const formattedHtmlPromises = Object.keys(formattedPagesContent).map(Number).sort((a, b) => a - b).map(async (pageNum) => `<h2>Page ${pageNum}</h2>\n${await marked.parse(formattedPagesContent[pageNum])}`);
    const allFormattedHtmlArray = await Promise.all(formattedHtmlPromises);
    const combinedHtml = allFormattedHtmlArray.join('<hr style="page-break-after: always; visibility: hidden;" />');
    if (combinedHtml.trim()) {
        try {
            const docStyles = `<style>table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid black; padding: 8px; }</style>`;
            const fullHtml = `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>Formatted Document</title>${docStyles}</head><body>${combinedHtml}</body></html>`;
            const fileBlob = await asBlob(fullHtml) as Blob;
            downloadFile(fileBlob, 'formatted_document.docx');
        } catch(err) {
            console.error("Error creating docx file:", err);
            showMessage(errorMessageEl, 'Could not create .docx file.', 'error');
        }
    } else showMessage(errorMessageEl, 'No formatted text to download.', 'error');
    downloadOptions.classList.add('hidden');
});
const pageNavListener = async (isNext: boolean, isExtracted: boolean) => {
    const content = isExtracted ? extractedPagesContent : formattedPagesContent;
    let currentPage = isExtracted ? extractedCurrentPage : formattedCurrentPage;
    const updateView = isExtracted ? updateExtractedView : updateFormattedView;
    const pageNums = Object.keys(content).map(Number).sort((a, b) => a - b);
    const currentIndex = pageNums.indexOf(currentPage);
    let newIndex = currentIndex;
    if (isNext && currentIndex < pageNums.length - 1) newIndex++;
    else if (!isNext && currentIndex > 0) newIndex--;
    if (newIndex !== currentIndex) {
        if (isExtracted) extractedCurrentPage = pageNums[newIndex];
        else formattedCurrentPage = pageNums[newIndex];
        await updateView();
    }
};
prevExtractedPageBtn.addEventListener('click', () => pageNavListener(false, true));
nextExtractedPageBtn.addEventListener('click', () => pageNavListener(true, true));
prevFormattedPageBtn.addEventListener('click', () => pageNavListener(false, false));
nextFormattedPageBtn.addEventListener('click', () => pageNavListener(true, false));
findBtnExtracted.addEventListener('click', () => findAndHighlight(extractedContentEl, findInputExtracted.value));
findInputExtracted.addEventListener('keydown', (e) => { if(e.key === 'Enter') { e.preventDefault(); findAndHighlight(extractedContentEl, findInputExtracted.value) }});
findBtnFormatted.addEventListener('click', () => findAndHighlight(formattedContentEl, findInputFormatted.value));
findInputFormatted.addEventListener('keydown', (e) => { if(e.key === 'Enter') { e.preventDefault(); findAndHighlight(formattedContentEl, findInputFormatted.value) }});
closeSummaryModalBtn.addEventListener('click', () => summaryModalEl.classList.add('hidden'));
copySummaryBtn.addEventListener('click', async () => {
    const summaryText = summaryContentEl.innerText;
    if (summaryText) {
        try {
            await navigator.clipboard.writeText(summaryText);
            copySummaryBtn.textContent = 'Copied!';
            setTimeout(() => { copySummaryBtn.textContent = 'Copy Summary'; }, 2000);
        } catch (err) { alert('Failed to copy summary text.'); }
    }
});
closeDownloadInfoModalBtn.addEventListener('click', () => downloadInfoModalEl.classList.add('hidden'));
if (followGateModalEl && followCheckboxEl && continueToAppBtn) {
    if (!localStorage.getItem('hasSeenFollowGate')) followGateModalEl.classList.remove('hidden');
    followCheckboxEl.addEventListener('change', () => {
        continueToAppBtn.disabled = !followCheckboxEl.checked;
        if (followCheckboxEl.checked) {
            continueToAppBtn.classList.replace('bg-gray-400', 'bg-blue-600');
            continueToAppBtn.classList.replace('cursor-not-allowed', 'cursor-pointer');
            continueToAppBtn.classList.add('hover:bg-blue-700');
        } else {
            continueToAppBtn.classList.replace('bg-blue-600', 'bg-gray-400');
            continueToAppBtn.classList.replace('cursor-pointer', 'cursor-not-allowed');
            continueToAppBtn.classList.remove('hover:bg-blue-700');
        }
    });
    continueToAppBtn.addEventListener('click', () => {
        if (!continueToAppBtn.disabled) {
            followGateModalEl.classList.add('hidden');
            localStorage.setItem('hasSeenFollowGate', 'true');
        }
    });
}

// --- New Feature Listeners (v2.0) ---
generateForewordBtn.addEventListener('click', generateForeword);
mergePdfsBtn.addEventListener('click', () => {
    clearAllMessages();
    pdfMergeModalEl.classList.remove('hidden');
});
bookCoverBtn.addEventListener('click', () => {
    clearAllMessages();
    updateCoverCanvas(); // Initialize canvas
    bookCoverModalEl.classList.remove('hidden');
});
closeWhatsNewModalBtn.addEventListener('click', () => whatsNewModalEl.classList.add('hidden'));
gotItWhatsNewBtn.addEventListener('click', () => whatsNewModalEl.classList.add('hidden'));
closeForewordModalBtn.addEventListener('click', () => forewordModalEl.classList.add('hidden'));
copyForewordBtn.addEventListener('click', async () => {
    const forewordText = forewordContentEl.innerText;
    if(forewordText) {
        await navigator.clipboard.writeText(forewordText);
        copyForewordBtn.textContent = 'Copied!';
        setTimeout(() => { copyForewordBtn.textContent = 'Copy Foreword' }, 2000);
    }
});
closePdfMergeModalBtn.addEventListener('click', () => pdfMergeModalEl.classList.add('hidden'));
pdfMergeInput.addEventListener('change', () => {
    executeMergeBtn.disabled = !pdfMergeInput.files || pdfMergeInput.files.length < 2;
});
executeMergeBtn.addEventListener('click', mergePdfs);
closeBookCoverModalBtn.addEventListener('click', () => bookCoverModalEl.classList.add('hidden'));
generateTitlesBtn.addEventListener('click', generateTitlesAndCaptions);
generateImageBtn.addEventListener('click', generateCoverImage);
coverReferenceImage.addEventListener('change', async (e) => {
    const file = (e.target as HTMLInputElement).files?.[0];
    if (file) coverState.referenceImage = await fileToBase64(file);
    else coverState.referenceImage = null;
});
coverAuthorInput.addEventListener('input', (e) => {
    coverState.author = (e.target as HTMLInputElement).value;
    updateCoverCanvas();
});
coverPriceInput.addEventListener('input', (e) => {
    coverState.price = (e.target as HTMLInputElement).value;
    updateCoverCanvas();
});
coverIsbnInput.addEventListener('input', (e) => {
    coverState.isbn = (e.target as HTMLInputElement).value;
    updateCoverCanvas();
});
downloadCoverBtn.addEventListener('click', assembleAndDownloadCover);

// --- Initial Setup ---
(async () => {
    await updateFormattedView();
    // What's New logic
    const lastSeenVersion = localStorage.getItem('lastSeenVersion');
    if(lastSeenVersion !== APP_VERSION) {
        whatsNewModalEl.classList.remove('hidden');
        localStorage.setItem('lastSeenVersion', APP_VERSION);
    }
})();