document.addEventListener("DOMContentLoaded", () => {

    let masterTemplateFile = null;
    let projectFiles = [];
    let template2File = null;

    const templateInput = document.getElementById("template-input");
    const templateNameDisplay = document.getElementById("template-name");
    const projectsInput = document.getElementById("projects-input");
    const sortableList = document.getElementById("sortable-list");
    const generateBtn = document.getElementById("generate-btn");
    const loadingOverlay = document.getElementById("loading-overlay");
    const modeSelect = document.getElementById("assembly-mode-select");
    const modeUnifiedDesc = document.getElementById("mode-unified-desc");
    const modeSplitDesc = document.getElementById("mode-split-desc");
    const template2Container = document.getElementById("template2-container");
    const template2Input = document.getElementById("template2-input");
    const template2NameDisplay = document.getElementById("template2-name");

    // Mode Switching
    modeSelect.addEventListener("change", (e) => {
        if (e.target.value === "split") {
            modeUnifiedDesc.classList.add("hidden");
            modeSplitDesc.classList.remove("hidden");
            template2Container.classList.remove("hidden");
            template2Container.classList.add("flex");
        } else {
            modeUnifiedDesc.classList.remove("hidden");
            modeSplitDesc.classList.add("hidden");
            template2Container.classList.add("hidden");
            template2Container.classList.remove("flex");
        }
    });

    template2Input.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            template2File = e.target.files[0];
            template2NameDisplay.textContent = template2File.name;
        }
    });

    let sortable = new Sortable(sortableList, { animation: 150, ghostClass: 'opacity-50' });

    templateInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            masterTemplateFile = e.target.files[0];
            templateNameDisplay.textContent = masterTemplateFile.name;
            templateNameDisplay.classList.add("text-white");
        }
    });

    projectsInput.addEventListener("change", (e) => {
        handleNewFiles(Array.from(e.target.files));
        projectsInput.value = "";
    });

    const projectsDropzone = document.getElementById("projects-dropzone");
    projectsDropzone.addEventListener("dragover", (e) => {
        e.preventDefault();
        projectsDropzone.style.borderColor = "#4ade80";
        projectsDropzone.style.background = "rgba(74, 222, 128, 0.1)";
    });
    projectsDropzone.addEventListener("dragleave", (e) => {
        e.preventDefault();
        projectsDropzone.style.borderColor = "";
        projectsDropzone.style.background = "";
    });
    projectsDropzone.addEventListener("drop", (e) => {
        e.preventDefault();
        projectsDropzone.style.borderColor = "";
        projectsDropzone.style.background = "";
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleNewFiles(Array.from(e.dataTransfer.files));
        }
    });

    function handleNewFiles(newFiles) {
        newFiles.forEach(file => {
            const uniqueId = "proj_" + Math.random().toString(36).substr(2, 9);
            projectFiles.push({ id: uniqueId, file: file });
            const li = document.createElement("li");
            li.className = "file-item rounded-md p-3 flex justify-between items-center cursor-move text-gray-300 shadow-sm transition hover:bg-gray-800";
            li.setAttribute("data-id", uniqueId);
            li.innerHTML = `
                <div class="flex items-center text-emerald-100">
                    <svg class="w-5 h-5 mr-3 text-emerald-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg>
                    <span class="truncate max-w-[200px] md:max-w-xs xl:max-w-md font-semibold text-sm">${file.name}</span>
                </div>
                <button class="delete-btn text-red-400 hover:text-red-300 p-1 z-10" data-id="${uniqueId}">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                </button>
            `;
            sortableList.appendChild(li);
        });
    }

    sortableList.addEventListener("click", (e) => {
        const btn = e.target.closest(".delete-btn");
        if (btn) {
            const idToRemove = btn.getAttribute("data-id");
            projectFiles = projectFiles.filter(item => item.id !== idToRemove);
            document.querySelector(`li[data-id="${idToRemove}"]`).remove();
        }
    });

    generateBtn.addEventListener("click", async () => {
        if (!masterTemplateFile) { alert("Por favor carga una Plantilla Maestra."); return; }
        const mode = modeSelect.value;
        if (mode === "split" && !template2File) { alert("Carga la Plantilla de Pago 2."); return; }
        if (projectFiles.length === 0) { alert("Carga al menos un archivo de proyecto."); return; }

        const formData = new FormData();
        formData.append("mode", mode);
        formData.append("masterTemplate", masterTemplateFile);
        if (mode === "split") formData.append("template2", template2File);
        formData.append("fontFamily", document.getElementById("font-family-select").value);
        formData.append("generateAI", document.getElementById("ai-toggle").checked);
        formData.append("exportFormat", document.getElementById("export-format-select").value);

        const currentOrderIds = Array.from(sortableList.children).map(li => li.getAttribute("data-id"));
        currentOrderIds.forEach(id => {
            const fileObj = projectFiles.find(item => item.id === id);
            if (fileObj) {
                formData.append("projectIds[]", id);
                formData.append(`projectFile_${id}`, fileObj.file);
            }
        });

        loadingOverlay.classList.remove("hidden");
        loadingOverlay.classList.add("flex");

        try {
            const response = await fetch('/api/generate_v5', { method: 'POST', body: formData });
            if (response.ok) {
                const data = await response.json();
                if (data.taskId) { await pollTaskProgress(data.taskId); }
                else if (data.error) { alert("Error: " + data.error); closeLoadingOverlay(); }
            } else {
                const errorData = await response.json();
                alert("Error: " + (errorData.error || "Fallo en servidor.")); closeLoadingOverlay();
            }
        } catch (error) { alert("Error de Red: " + error.message); closeLoadingOverlay(); }
    });

    async function pollTaskProgress(taskId) {
        const progressBar = document.getElementById("progress-bar-fill");
        const progressPct = document.getElementById("progress-percentage");
        const progressText = document.getElementById("progress-status-text");
        const aiErrorContainer = document.getElementById("error-ai-container");
        const aiErrorTextSpan = document.querySelector("#error-ai-text span");

        const pollInterval = setInterval(async () => {
            try {
                const res = await fetch('/api/progress/' + taskId);
                const data = await res.json();
                progressBar.style.width = data.progress + "%";
                progressPct.textContent = data.progress + "%";
                if (data.status) progressText.textContent = data.status;
                if (data.error_ai) {
                    aiErrorTextSpan.textContent = "Rebote API Grok (" + data.error_ai + ")";
                    aiErrorContainer.classList.remove("hidden");
                }
                if (data.completed) {
                    clearInterval(pollInterval);
                    closeLoadingOverlay();
                    if (data.success) {
                        const resultsPanel = document.getElementById("results-panel");
                        const aiContainer = document.getElementById("ai-options-container");
                        resultsPanel.classList.remove("hidden");

                        if (data.downloadUrls) {
                            showBtn("download-unified-word-btn", data.downloadUrls.docx);
                            showBtn("download-unified-pdf-btn", data.downloadUrls.pdf);
                            showBtn("download-csv-btn", data.downloadUrls.csv);
                            showBtn("download-links-docx-btn", data.downloadUrls.links_docx);
                        }
                        if (data.aiDownloadUrls || data.aiPredictiveUrls) {
                            aiContainer.classList.remove("hidden");
                            aiContainer.classList.add("flex");
                            if (data.aiDownloadUrls && data.aiDownloadUrls.pdf) showBtn("download-ai-btn", data.aiDownloadUrls.pdf);
                            if (data.aiPredictiveUrls && data.aiPredictiveUrls.pdf) showBtn("download-predictive-btn", data.aiPredictiveUrls.pdf);
                        }
                        let primaryDownload = data.downloadUrls?.docx || data.downloadUrls?.pdf;
                        if (primaryDownload) window.location.href = primaryDownload;
                        generateBtn.textContent = "✓ Finalizado. Generar Nuevo";
                    } else {
                        alert("Error Crítico: " + data.error);
                    }
                }
            } catch (err) { console.error("Polling error", err); }
        }, 800);
    }

    function showBtn(id, url) {
        const btn = document.getElementById(id);
        if (url && btn) { btn.href = url; btn.classList.remove("hidden"); btn.classList.add("flex"); }
        else if (btn) { btn.classList.add("hidden"); btn.classList.remove("flex"); }
    }

    function closeLoadingOverlay() {
        loadingOverlay.classList.remove("flex");
        loadingOverlay.classList.add("hidden");
    }
});
