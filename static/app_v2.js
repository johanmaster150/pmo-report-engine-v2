document.addEventListener("DOMContentLoaded", () => {

    let masterTemplateFile = null;
    let projectFiles = [];

    const templateInput = document.getElementById("template-input");
    const templateNameDisplay = document.getElementById("template-name");

    const projectsInput = document.getElementById("projects-input");
    const sortableList = document.getElementById("sortable-list");

    const generateBtn = document.getElementById("generate-btn");
    const loadingOverlay = document.getElementById("loading-overlay");

    // Phase 5 Routing Elements
    let template2File = null;
    const modeSelect = document.getElementById("assembly-mode-select");
    const modeUnifiedDesc = document.getElementById("mode-unified-desc");
    const modeSplitDesc = document.getElementById("mode-split-desc");
    const template2Container = document.getElementById("template2-container");
    const template2Input = document.getElementById("template2-input");
    const template2NameDisplay = document.getElementById("template2-name");

    // Handle Mode Switching
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

    // Handle Template 2 Upload
    template2Input.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            template2File = e.target.files[0];
            template2NameDisplay.textContent = template2File.name;
            template2NameDisplay.classList.remove("text-blue-400");
            template2NameDisplay.classList.add("text-white");
        }
    });

    // Initialize SortableJS
    let sortable = new Sortable(sortableList, {
        animation: 150,
        ghostClass: 'opacity-50', // Class applied while dragging
    });

    // Handle Master Template Upload
    templateInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            masterTemplateFile = e.target.files[0];
            templateNameDisplay.textContent = masterTemplateFile.name;
            templateNameDisplay.classList.remove("text-indigo-400");
            templateNameDisplay.classList.add("text-white");
        }
    });

    // Handle Project Files Upload (Multiple) via Click
    projectsInput.addEventListener("change", (e) => {
        handleNewFiles(Array.from(e.target.files));
        projectsInput.value = "";
    });

    // Native Drag and Drop listeners to prevent browser redirecting
    const projectsDropzone = document.getElementById("projects-dropzone");

    projectsDropzone.addEventListener("dragover", (e) => {
        e.preventDefault();
        projectsDropzone.style.borderColor = "#4ade80"; // focus emerald
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
            // Give each file a unique ID to track order
            const uniqueId = "proj_" + Math.random().toString(36).substr(2, 9);
            projectFiles.push({ id: uniqueId, file: file });

            // Render UI Element
            const li = document.createElement("li");
            li.className = "file-item rounded-md p-3 flex justify-between items-center cursor-move text-gray-300 shadow-sm transition hover:bg-gray-800";
            li.setAttribute("data-id", uniqueId);

            li.innerHTML = `
                <div class="flex items-center text-emerald-100">
                    <svg class="w-5 h-5 mr-3 text-emerald-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg>
                    <span class="truncate max-w-[200px] md:max-w-xs xl:max-w-md font-semibold text-sm tracking-wide">${file.name}</span>
                </div>
                <!-- Delete Button -->
                <button class="delete-btn text-red-400 hover:text-red-300 transition-colors p-1 group z-10" data-id="${uniqueId}">
                    <svg class="w-5 h-5 group-hover:scale-110 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                </button>
            `;
            sortableList.appendChild(li);
        });
    }

    // Handle Item Deletion
    sortableList.addEventListener("click", (e) => {
        const btn = e.target.closest(".delete-btn");
        if (btn) {
            const idToRemove = btn.getAttribute("data-id");
            // Remove from array
            projectFiles = projectFiles.filter(item => item.id !== idToRemove);
            // Remove from DOM
            document.querySelector(`li[data-id="${idToRemove}"]`).remove();
        }
    });

    // Form Submission (Generate Button)
    generateBtn.addEventListener("click", async () => {

        if (!masterTemplateFile) {
            alert("⚠️ Por favor carga una Plantilla Maestra primero.");
            return;
        }

        const mode = modeSelect.value;
        if (mode === "split" && !template2File) {
            alert("⚠️ En el Modo Pagos Divididos necesitas cargar también la Plantilla de Pago 2.");
            return;
        }

        if (projectFiles.length === 0) {
            alert("⚠️ Por favor carga al menos un archivo de proyecto.");
            return;
        }

        // Prepare Form Data
        const formData = new FormData();
        formData.append("mode", mode);
        formData.append("masterTemplate", masterTemplateFile);
        if (mode === "split") {
            formData.append("template2", template2File);
        }
        formData.append("fontFamily", document.getElementById("font-family-select").value);
        formData.append("generateAI", document.getElementById("ai-toggle").checked);
        formData.append("exportFormat", document.getElementById("export-format-select").value);
        formData.append("generalFontSize", document.getElementById("general-font-size").value);
        formData.append("titleFontSize", document.getElementById("title-font-size").value);
        formData.append("subtitleFontSize", document.getElementById("subtitle-font-size").value);
        formData.append("tableFontSize", document.getElementById("table-font-size").value);
        formData.append("captionFontSize", document.getElementById("caption-font-size").value);

        // Get the CURRENT order from the DOM (Sortable modifies the DOM directly)
        const currentOrderIds = Array.from(sortableList.children).map(li => li.getAttribute("data-id"));

        currentOrderIds.forEach(id => {
            // Find the actual file object for this ID
            const fileObj = projectFiles.find(item => item.id === id);
            if (fileObj) {
                // Send the ID in the array, and the file separately keyed by ID
                formData.append("projectIds[]", id);
                formData.append(`projectFile_${id}`, fileObj.file);
            }
        });

        // Show UI Loading State
        loadingOverlay.classList.remove("hidden");
        loadingOverlay.classList.add("flex");

        try {
            const response = await fetch('/api/generate_v2', {
                method: 'POST',
                body: formData
            });

            if (response.ok) {
                const data = await response.json();
                if (data.taskId) {
                    // Start Polling
                    await pollTaskProgress(data.taskId);
                } else if (data.error) {
                    alert("❌ Error Server: " + data.error);
                    closeLoadingOverlay();
                }
            } else {
                const errorData = await response.json();
                alert("❌ Error: " + (errorData.error || "Falló la generación en el servidor."));
                closeLoadingOverlay();
            }
        } catch (error) {
            alert("❌ Error de Red o Servidor inactivo: " + error.message);
            closeLoadingOverlay();
        }
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

                // Update UI visually
                progressBar.style.width = data.progress + "%";
                progressPct.textContent = data.progress + "%";
                if (data.status) progressText.textContent = data.status;

                if (data.error_ai) {
                    aiErrorTextSpan.textContent = "Rebote API OpenRouter (" + data.error_ai + ")";
                    aiErrorContainer.classList.remove("hidden");
                }

                if (data.completed) {
                    clearInterval(pollInterval);
                    closeLoadingOverlay();

                    if (data.success) {
                        const resultsPanel = document.getElementById("results-panel");
                        const unifiedContainer = document.getElementById("unified-downloads-container");
                        const aiContainer = document.getElementById("ai-options-container");

                        const unifiedWordBtn = document.getElementById("download-unified-word-btn");
                        const unifiedPdfBtn = document.getElementById("download-unified-pdf-btn");
                        const aiBtn = document.getElementById("download-ai-btn");

                        resultsPanel.classList.remove("hidden");

                        // Handle Unified Downloads
                        if (data.downloadUrls) {
                            if (data.downloadUrls.docx) {
                                unifiedWordBtn.href = data.downloadUrls.docx;
                                unifiedWordBtn.classList.remove("hidden");
                                unifiedWordBtn.classList.add("flex");
                            } else {
                                unifiedWordBtn.classList.add("hidden");
                                unifiedWordBtn.classList.remove("flex");
                            }

                            if (data.downloadUrls.pdf) {
                                unifiedPdfBtn.href = data.downloadUrls.pdf;
                                unifiedPdfBtn.classList.remove("hidden");
                                unifiedPdfBtn.classList.add("flex");
                            } else {
                                unifiedPdfBtn.classList.add("hidden");
                                unifiedPdfBtn.classList.remove("flex");
                            }

                            // CSV de Links y Evidencias
                            const csvBtn = document.getElementById("download-csv-btn");
                            if (data.downloadUrls.csv) {
                                csvBtn.href = data.downloadUrls.csv;
                                csvBtn.classList.remove("hidden");
                                csvBtn.classList.add("flex");
                            } else {
                                csvBtn.classList.add("hidden");
                                csvBtn.classList.remove("flex");
                            }

                            // DOCX de Links y Evidencias
                            const linksDocxBtn = document.getElementById("download-links-docx-btn");
                            if (data.downloadUrls.links_docx) {
                                linksDocxBtn.href = data.downloadUrls.links_docx;
                                linksDocxBtn.classList.remove("hidden");
                                linksDocxBtn.classList.add("flex");
                            } else {
                                linksDocxBtn.classList.add("hidden");
                                linksDocxBtn.classList.remove("flex");
                            }
                        }

                        // Handle AI Downloads (Audit + Predictive in same container)
                        if (data.aiDownloadUrls || data.aiPredictiveUrls) {
                            aiContainer.classList.remove("hidden");
                            aiContainer.classList.add("flex");

                            if (data.aiDownloadUrls && data.aiDownloadUrls.pdf) {
                                aiBtn.href = data.aiDownloadUrls.pdf;
                                aiBtn.classList.remove("hidden");
                                aiBtn.classList.add("flex");
                            }

                            const predictiveBtn = document.getElementById("download-predictive-btn");
                            if (data.aiPredictiveUrls && data.aiPredictiveUrls.pdf) {
                                predictiveBtn.href = data.aiPredictiveUrls.pdf;
                                predictiveBtn.classList.remove("hidden");
                                predictiveBtn.classList.add("flex");
                            }
                        } else {
                            aiContainer.classList.add("hidden");
                            aiContainer.classList.remove("flex");
                        }

                        // Only auto-download the word doc if it exists, otherwise pdf
                        let primaryDownload = data.downloadUrls?.docx || data.downloadUrls?.pdf;
                        if (primaryDownload) window.location.href = primaryDownload;

                        generateBtn.textContent = "✓ Finalizado. Generar Nuevo";
                    } else {
                        alert("❌ Error Crítico: " + data.error);
                    }
                }
            } catch (err) {
                console.error("Polling error", err);
            }
        }, 800);
    }

    // Handlers for PDF Modal
    const closePdfModalBtn = document.getElementById("close-pdf-modal");
    if (closePdfModalBtn) {
        closePdfModalBtn.addEventListener('click', () => {
            const modal = document.getElementById("pdf-modal");
            const iframe = document.getElementById("pdf-iframe");
            modal.classList.add("hidden");
            modal.classList.remove("flex");
            // Clear iframe to stop memory consumption when hidden
            setTimeout(() => { iframe.src = ""; }, 300);
        });
    }

    function closeLoadingOverlay() {
        loadingOverlay.classList.remove("flex");
        loadingOverlay.classList.add("hidden");
    }
});
