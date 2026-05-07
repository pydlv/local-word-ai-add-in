import React from "react";

const SETTINGS_KEY = "offline-writer-settings";
const ACTIVE_SELECTION_BOOKMARK = "_LocalWriterSelection";
const LEGACY_CONTENT_CONTROL_TAGS = ["local-writer-last-action", "local-writer-active-selection"];

interface LocalModel {
    id: string;
}

type WriterMode = "insert" | "replace";

interface Settings {
    endpoint: string;
    model: string;
    mode: WriterMode;
    maxContextChars: number;
    maxOutputTokens: number;
    systemPrompts: Record<WriterMode, string>;
}

interface HomeState extends Settings {
    prompt: string;
    models: LocalModel[];
    isLoading: boolean;
    status: string;
    showSettings: boolean;
}

interface ContextResult {
    context: string;
}

export enum Page {
    Home = "Home",
    GeneratedPage = "GeneratedPage",
    Chat = "Chat",
}

const defaultSystemPrompts: Record<WriterMode, string> = {
    insert:
        "You are an offline writing assistant inside Microsoft Word. Use the document context and the cursor marker to write one polished paragraph for the cursor location. Return only the paragraph text.",
    replace:
        [
            "You are an offline span editor inside Microsoft Word.",
            "Rewrite only the selected text according to the user's prompt.",
            "The text before and after the selection is already in the document and must not be repeated.",
            "Return only the replacement span that can be pasted exactly between BEFORE_SELECTION and AFTER_SELECTION.",
            "If the selection is a sentence fragment, return a sentence fragment. Do not expand it into a full sentence or paragraph.",
        ].join(" "),
};

const defaultSettings: Settings = {
    endpoint: getDefaultEndpoint(),
    model: "",
    mode: "insert",
    maxContextChars: 4000,
    maxOutputTokens: 180,
    systemPrompts: { ...defaultSystemPrompts },
};

function isLocalAddInHost() {
    return window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";
}

function getDefaultEndpoint() {
    return isLocalAddInHost() ? "/local-ai" : "http://127.0.0.1:1234";
}

export default class Home extends React.Component<Record<string, never>, HomeState> {
    state: HomeState = {
        ...this.loadSettings(),
        prompt: "",
        models: [],
        isLoading: false,
        status: "",
        showSettings: false,
    };

    componentDidMount() {
        this.refreshModels();
        this.cleanupLegacyContentControls();
    }

    loadSettings(): Settings {
        try {
            const saved = window.localStorage.getItem(SETTINGS_KEY);
            const parsedSettings = saved ? JSON.parse(saved) : {};
            const settings = {
                ...defaultSettings,
                ...parsedSettings,
                systemPrompts: {
                    ...defaultSystemPrompts,
                    ...(parsedSettings.systemPrompts || {}),
                },
            };

            if (isLocalAddInHost() && (settings.endpoint === "http://127.0.0.1:1234" || settings.endpoint === "http://localhost:1234")) {
                settings.endpoint = "/local-ai";
            }

            return settings;
        } catch {
            return defaultSettings;
        }
    }

    saveSettings(nextState: Partial<Settings>) {
        const settings = {
            endpoint: nextState.endpoint ?? this.state.endpoint,
            model: nextState.model ?? this.state.model,
            mode: nextState.mode ?? this.state.mode,
            maxContextChars: nextState.maxContextChars ?? this.state.maxContextChars,
            maxOutputTokens: nextState.maxOutputTokens ?? this.state.maxOutputTokens,
            systemPrompts: nextState.systemPrompts ?? this.state.systemPrompts,
        };

        window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
    }

    setSetting<K extends Exclude<keyof Settings, "systemPrompts">>(key: K, value: Settings[K]) {
        this.setState({ [key]: value } as any, () => this.saveSettings({ [key]: value } as Partial<Settings>));
    }

    setSystemPrompt(mode: WriterMode, value: string) {
        const systemPrompts = { ...this.state.systemPrompts, [mode]: value };

        this.setState({ systemPrompts }, () => this.saveSettings({ systemPrompts }));
    }

    resetSystemPrompt(mode: WriterMode) {
        this.setSystemPrompt(mode, defaultSystemPrompts[mode]);
    }

    normalizeEndpoint() {
        const endpoint = this.state.endpoint.replace(/\/+$/, "");

        if (endpoint === "http://127.0.0.1:1234" || endpoint === "http://localhost:1234") {
            return isLocalAddInHost() ? "/local-ai" : "http://127.0.0.1:1234";
        }

        return endpoint === "/local-ai" && !isLocalAddInHost() ? "http://127.0.0.1:1234" : endpoint;
    }

    async refreshModels() {
        this.setState({ status: "Checking local models..." });

        try {
            const response = await fetch(`${this.normalizeEndpoint()}/v1/models`);

            if (!response.ok) {
                throw new Error(`Model list failed: ${response.status}`);
            }

            const data = await response.json();
            const models = this.parseModels(data);
            const model = this.state.model || models[0]?.id || "";

            this.setState({ models, model, status: models.length ? "Local model ready." : "No models returned." }, () =>
                this.saveSettings({ model })
            );
        } catch (error) {
            this.setState({ models: [], status: this.formatError(error) });
        }
    }

    parseModels(data: any): LocalModel[] {
        const rawModels = Array.isArray(data?.data)
            ? data.data
            : Array.isArray(data)
            ? data
            : data?.models && typeof data.models === "object"
            ? Object.keys(data.models)
            : [];

        return rawModels
            .map((model) => {
                if (typeof model === "string") {
                    return { id: model };
                }

                const id = model?.id || model?.name || model?.model;

                return id ? { id: String(id) } : null;
            })
            .filter((model): model is LocalModel => model !== null);
    }

    async getContextAroundCursor(): Promise<ContextResult> {
        return Word.run(async (ctx) => {
            const selection = ctx.document.getSelection();
            const cursor = selection.getRange(Word.RangeLocation.start);
            const bodyRange = selection.parentBody.getRange(Word.RangeLocation.content);

            cursor.load("start");
            bodyRange.load(["start", "text"]);
            await ctx.sync();

            const bodyText = bodyRange.text || "";
            const cursorIndex = Math.max(0, Math.min(bodyText.length, cursor.start - bodyRange.start));
            const before = bodyText.slice(0, cursorIndex).replace(/\s+$/g, "");
            const after = bodyText.slice(cursorIndex).replace(/^\s+/g, "");
            const window = this.createRollingWindow(before, after);

            return { context: [window.before, "[[CURSOR]]", window.after].filter(Boolean).join("\n") };
        });
    }

    async getContextAroundSelection(): Promise<ContextResult> {
        return Word.run(async (ctx) => {
            ctx.document.deleteBookmark(ACTIVE_SELECTION_BOOKMARK);

            const selection = ctx.document.getSelection();
            const bodyRange = selection.parentBody.getRange(Word.RangeLocation.content);

            selection.load(["start", "end", "text"]);
            bodyRange.load(["start", "text"]);
            await ctx.sync();

            const selectedText = selection.text || "";

            if (!selectedText.trim()) {
                throw new Error("Select text to replace first.");
            }

            const bodyText = bodyRange.text || "";
            const selectionStart = Math.max(0, Math.min(bodyText.length, selection.start - bodyRange.start));
            const selectionEnd = Math.max(selectionStart, Math.min(bodyText.length, selection.end - bodyRange.start));
            const before = bodyText.slice(0, selectionStart).replace(/\s+$/g, "");
            const after = bodyText.slice(selectionEnd).replace(/^\s+/g, "");
            const window = this.createRollingWindow(before, after);

            selection.insertBookmark(ACTIVE_SELECTION_BOOKMARK);
            await ctx.sync();

            return {
                context: [
                    "BEFORE_SELECTION:",
                    window.before || "(none)",
                    "",
                    "SELECTED_TEXT:",
                    selectedText,
                    "",
                    "AFTER_SELECTION:",
                    window.after || "(none)",
                ].join("\n"),
            };
        });
    }

    createRollingWindow(beforeAll: string, afterAll: string) {
        const max = Math.max(0, Number(this.state.maxContextChars) || 0);

        if (max === 0) {
            return { before: "", after: "" };
        }

        const beforeTarget = Math.floor(max / 2);
        const afterTarget = max - beforeTarget;
        let beforeLength = Math.min(beforeAll.length, beforeTarget);
        let afterLength = Math.min(afterAll.length, afterTarget);
        let unused = max - beforeLength - afterLength;

        if (unused > 0 && beforeLength < beforeAll.length) {
            const extra = Math.min(unused, beforeAll.length - beforeLength);
            beforeLength += extra;
            unused -= extra;
        }

        if (unused > 0 && afterLength < afterAll.length) {
            afterLength += Math.min(unused, afterAll.length - afterLength);
        }

        return {
            before: beforeAll.slice(beforeAll.length - beforeLength),
            after: afterAll.slice(0, afterLength),
        };
    }

    buildUserPrompt(context: string, mode: WriterMode = this.state.mode) {
        const parts = [
            "Document context:",
            context || (mode === "insert" ? "[[CURSOR]]" : "BEFORE_SELECTION:\n(none)\n\nSELECTED_TEXT:\n\nAFTER_SELECTION:\n(none)"),
            "",
            "User prompt:",
            this.state.prompt.trim(),
        ];

        if (mode === "replace") {
            parts.push(
                "",
                "Replacement rules:",
                "- Return only the new text for SELECTED_TEXT.",
                "- Do not include BEFORE_SELECTION or AFTER_SELECTION.",
                "- Do not duplicate words that are already adjacent to the selection.",
                "- Preserve sentence flow at both boundaries."
            );
        }

        return parts.join("\n");
    }

    async generateText(context: string, mode: WriterMode = this.state.mode): Promise<string> {
        const endpoint = this.normalizeEndpoint();
        const model = this.state.model.trim();
        const userPrompt = this.buildUserPrompt(context, mode);

        if (!model) {
            throw new Error("Select a local model first.");
        }

        const systemPrompt = this.state.systemPrompts[mode].trim() || defaultSystemPrompts[mode];
        const temperature = mode === "replace" ? 0.2 : 0.7;

        try {
            const response = await this.postJson(`${endpoint}/v1/chat/completions`, {
                model,
                messages: [
                    { role: "system", content: systemPrompt },
                    { role: "user", content: userPrompt },
                ],
                max_tokens: this.state.maxOutputTokens,
                temperature,
            });

            return this.extractChatText(response);
        } catch {
            try {
                const response = await this.postJson(`${endpoint}/v1/responses`, {
                    model,
                    input: [
                        { role: "system", content: systemPrompt },
                        { role: "user", content: userPrompt },
                    ],
                    max_output_tokens: this.state.maxOutputTokens,
                    temperature,
                });

                return this.extractResponsesText(response);
            } catch {
                const response = await this.postJson(`${endpoint}/v1/completions`, {
                    model,
                    prompt: `${systemPrompt}\n\n${userPrompt}`,
                    max_tokens: this.state.maxOutputTokens,
                    temperature,
                });

                return this.extractCompletionText(response);
            }
        }
    }

    async postJson(url: string, body: unknown) {
        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data?.error?.message || data?.message || `${response.status} ${response.statusText}`);
        }

        return data;
    }

    extractResponsesText(data: any): string {
        if (typeof data.output_text === "string") {
            return data.output_text.trim();
        }

        const text = (data.output || [])
            .reduce((parts, item) => parts.concat(item.content || []), [])
            .map((content) => (typeof content === "string" ? content : content.text || content.output_text || ""))
            .join("")
            .trim();

        if (!text) {
            throw new Error("The local model returned no text.");
        }

        return text;
    }

    extractCompletionText(data: any): string {
        const text = data.choices?.[0]?.text?.trim();

        if (!text) {
            throw new Error("The local model returned no text.");
        }

        return text;
    }

    extractChatText(data: any): string {
        const text = data.choices?.[0]?.message?.content?.trim();

        if (!text) {
            throw new Error("The local model returned no text.");
        }

        return text;
    }

    async cleanupLegacyContentControls() {
        await Word.run(async (ctx) => {
            const collections = LEGACY_CONTENT_CONTROL_TAGS.map((tag) => {
                const collection = ctx.document.contentControls.getByTag(tag);

                collection.load("items");
                return collection;
            });

            await ctx.sync();
            collections.forEach((collection) => collection.items.forEach((control) => control.delete(true)));
            await ctx.sync();
        }).catch(() => undefined);
    }

    async insertAtCursor(text: string) {
        await Word.run(async (ctx) => {
            ctx.document.deleteBookmark(ACTIVE_SELECTION_BOOKMARK);

            const selection = ctx.document.getSelection();
            const inserted = selection.insertText(text.trim() + "\n", Word.InsertLocation.replace);

            inserted.select(Word.SelectionMode.end);
            await ctx.sync();
        });
    }

    async replaceSelection(text: string) {
        await Word.run(async (ctx) => {
            const selectionRange = ctx.document.getBookmarkRangeOrNullObject(ACTIVE_SELECTION_BOOKMARK);

            selectionRange.load("isNullObject");
            await ctx.sync();

            if (selectionRange.isNullObject) {
                throw new Error("Could not find the selected text to replace.");
            }

            const inserted = selectionRange.insertText(text.trim(), Word.InsertLocation.replace);

            ctx.document.deleteBookmark(ACTIVE_SELECTION_BOOKMARK);
            inserted.select();
            await ctx.sync();
        });
    }

    formatError(error: unknown) {
        return error instanceof Error ? error.message : String(error);
    }

    handleSubmit = async () => {
        if (!this.state.prompt.trim()) {
            this.setState({ status: "Enter a prompt first." });
            return;
        }

        this.setState({ isLoading: true, status: "Reading cursor context..." });

        try {
            const contextResult =
                this.state.mode === "replace" ? await this.getContextAroundSelection() : await this.getContextAroundCursor();

            this.setState({ status: "Generating locally..." });

            const text = await this.generateText(contextResult.context);

            this.setState({ status: this.state.mode === "replace" ? "Replacing..." : "Inserting..." });

            if (this.state.mode === "replace") {
                await this.replaceSelection(text);
                this.setState({ status: "Replaced." });
            } else {
                await this.insertAtCursor(text);
                this.setState({ status: "Inserted." });
            }
        } catch (error) {
            await Word.run(async (ctx) => {
                ctx.document.deleteBookmark(ACTIVE_SELECTION_BOOKMARK);
                await ctx.sync();
            }).catch(() => undefined);
            this.setState({ status: this.formatError(error) });
        } finally {
            this.setState({ isLoading: false });
        }
    };

    render() {
        const currentSystemPrompt = this.state.systemPrompts[this.state.mode] || defaultSystemPrompts[this.state.mode];

        return (
            <main className="offlinePane">
                <header className="paneHeader">
                    <h1>Local Writer</h1>
                    <button type="button" onClick={() => this.setState({ showSettings: !this.state.showSettings })}>
                        {this.state.showSettings ? "Hide Settings" : "Settings"}
                    </button>
                </header>

                <div className="modeGroup" role="group" aria-label="Writing mode">
                    <button
                        type="button"
                        className={this.state.mode === "insert" ? "modeButton active" : "modeButton"}
                        onClick={() => this.setSetting("mode", "insert")}
                    >
                        Insert
                    </button>
                    <button
                        type="button"
                        className={this.state.mode === "replace" ? "modeButton active" : "modeButton"}
                        onClick={() => this.setSetting("mode", "replace")}
                    >
                        Replace
                    </button>
                </div>

                {this.state.showSettings && (
                    <section className="settingsGrid">
                        <label>
                            Endpoint
                            <input
                                value={this.state.endpoint}
                                onChange={(event) => this.setSetting("endpoint", event.target.value)}
                                onBlur={() => this.refreshModels()}
                            />
                        </label>

                        <label>
                            Model
                            <input
                                list="local-models"
                                value={this.state.model}
                                onChange={(event) => this.setSetting("model", event.target.value)}
                                placeholder="Type or select a model id"
                            />
                            <datalist id="local-models">
                                {this.state.models.map((model) => (
                                    <option key={model.id} value={model.id} />
                                ))}
                            </datalist>
                        </label>

                        <label>
                            Max context chars
                            <input
                                type="number"
                                min="0"
                                step="500"
                                value={this.state.maxContextChars}
                                onChange={(event) => this.setSetting("maxContextChars", Number(event.target.value))}
                            />
                        </label>

                        <label>
                            Max output tokens
                            <input
                                type="number"
                                min="32"
                                step="16"
                                value={this.state.maxOutputTokens}
                                onChange={(event) => this.setSetting("maxOutputTokens", Number(event.target.value))}
                            />
                        </label>

                        <label className="systemPromptBox">
                            <span className="systemPromptHeader">
                                System prompt
                                <button
                                    type="button"
                                    className="resetPromptButton"
                                    onClick={() => this.resetSystemPrompt(this.state.mode)}
                                    disabled={this.state.isLoading || currentSystemPrompt === defaultSystemPrompts[this.state.mode]}
                                >
                                    Reset
                                </button>
                            </span>
                            <textarea
                                value={currentSystemPrompt}
                                onChange={(event) => this.setSystemPrompt(this.state.mode, event.target.value)}
                            />
                        </label>

                        <button type="button" onClick={() => this.refreshModels()} disabled={this.state.isLoading}>
                            Refresh Models
                        </button>
                    </section>
                )}

                <label className="promptBox">
                    Prompt
                    <textarea
                        value={this.state.prompt}
                        onChange={(event) => this.setState({ prompt: event.target.value })}
                        placeholder="Write the next paragraph..."
                    />
                </label>

                <button className="primaryAction" type="button" onClick={this.handleSubmit} disabled={this.state.isLoading}>
                    {this.state.isLoading ? "Working..." : this.state.mode === "replace" ? "Replace Selection" : "Insert Paragraph"}
                </button>

                <div className="status" role="status">
                    {this.state.status}
                </div>
            </main>
        );
    }
}
