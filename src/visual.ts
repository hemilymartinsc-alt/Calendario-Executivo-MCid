import powerbi from "powerbi-visuals-api";
import RuntimeVisual from "./runtime";
import MCID_LOGO_DATA_URI from "./logo";
import "./../style/visual.less";
import "./../style/v3-opcao2.less";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

const MONTHS = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const CAT_ORDER = ["Entrega","Assinatura","Emissão de O.S.","Visita","Institucional","Outros"];
const COLORS: {[key:string]:string} = {
    "Entrega":"#0C9A43",
    "Assinatura":"#DF1717",
    "Emissão de O.S.":"#7A1FA2",
    "Visita":"#F5B400",
    "Institucional":"#0B5CC7",
    "Outros":"#90959D"
};

const LEGEND_LABELS: {[key:string]:string} = {
    "Entrega":"Entregas",
    "Assinatura":"Assinaturas",
    "Emissão de O.S.":"Emissão de O.S.",
    "Visita":"Visitas",
    "Institucional":"Institucionais",
    "Outros":"Outros"
};


function esc(value:any): string {
    return String(value == null ? "" : value).replace(/[&<>\"']/g, (c:string) => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c]));
}

function norm(value:any): string {
    return (value == null ? "" : String(value)).normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
}

function hasValue(value:any): boolean {
    if(value === null || value === undefined) return false;
    const s = String(value).trim();
    return s !== "" && s !== "–" && s !== "-";
}

function isYes(value:any): boolean {
    const s = norm(value);
    return s === "sim" || s === "s" || s === "yes" || s === "true" || s === "1";
}

function pad2(value:number): string {
    return value < 10 ? "0" + value : String(value);
}

function fmtDate(d:Date): string {
    if(!d) return "";
    return pad2(d.getDate()) + "/" + pad2(d.getMonth()+1) + "/" + d.getFullYear();
}

function sameDate(a:Date,b:Date): boolean {
    return !!(a && b && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate());
}

function errorMessage(error:any): string {
    try {
        if(error && error.message) return String(error.message).slice(0,360);
        if(error !== null && error !== undefined) return String(error).slice(0,360);
    } catch(_e) {}
    return "Erro não identificado.";
}

/**
 * V3 / Opção 2.
 * A classe continua herdando diretamente o runtime 1.2.1 que já funciona no Power BI.
 * As mudanças são aplicadas como uma camada de evolução, preservando o bootstrap original.
 */
export class Visual extends (RuntimeVisual as any) implements IVisual {
    constructor(options: VisualConstructorOptions) {
        super(options);
        const self:any = this;
        if(!self.state) self.state = {};
        if(self.state.expandedEventId === undefined) self.state.expandedEventId = null;
    }

    public showLanding(): void {
        const self:any = this;
        if(!self.root) return;
        self.root.innerHTML = '<div class="mcid-landing"><div class="mcid-landing-box"><h2>Calendário Executivo MCid</h2><p>O visual foi importado corretamente. Para ativá-lo, associe os campos da tabela <b>Base</b> aos campos do visual.</p><div class="mcid-map-table"><div><b>Id Evento</b> → Id Evento</div><div><b>Data Específica</b> → Data Específica (se houver)</div><div><b>Data Quinzena</b> → Data Quinzena</div><div><b>Ano do Evento</b> → Ano do Evento</div><div><b>Tipo Data</b> → Tipo Data</div><div><b>Quinzena Nº</b> → Quinzena Nº</div><div><b>Categoria Calendário</b> → Categoria Calendário</div><div><b>Tipologia</b> → Tipologia do Evento</div><div><b>Empreendimento</b> → Nome do Empreendimento / Medida / Ação</div><div><b>Município</b> → Município</div><div><b>UF</b> → Sigla Estado</div><div><b>Secretaria</b> → Unidade responsável</div><div><b>Fonte</b> → Subfonte</div><div><b>Executor</b> → Executor do Empreendimento</div><div><b>Minha Casa, Minha Vida</b> → Minha casa minha vida</div><div><b>Novo PAC</b> → Novo PAC (sim/não)</div><div><b>UH</b> → UH</div><div><b>Observações</b> → Observações</div></div><div class="mcid-warn">O visual não cria datas artificiais: Data Específica vai ao dia exato; Data Quinzena fica nas previsões por quinzena; registros sem ambas permanecem em Sem Data.</div></div></div>';
    }

    private showRuntimeError(stage:string, error:any): void {
        const self:any = this;
        try {
            if(!self.root && self.target) {
                self.root = document.createElement("div");
                self.root.className = "mcid-calendar-root";
                self.target.appendChild(self.root);
            }
            if(!self.root) return;
            self.root.innerHTML = '<div class="mcid-runtime-error"><div class="mcid-runtime-error-box"><div class="mcid-runtime-error-title">Não foi possível exibir o Calendário Executivo MCid</div><div class="mcid-runtime-error-stage">Etapa: ' + esc(stage) + '</div><div class="mcid-runtime-error-message">' + esc(errorMessage(error)) + '</div><div class="mcid-runtime-error-help">O visual permanece carregado. Revise os campos associados e atualize o visual.</div></div></div>';
        } catch(_e) {}
    }

    public update(options: VisualUpdateOptions): void {
        const self:any = this;
        try {
            // O update original é mantido porque é o caminho já validado no Power BI.
            super.update(options);

            const dv:any = options && options.dataViews && options.dataViews[0];
            if(!dv || !dv.table || !dv.table.rows || !dv.table.columns || !self.events) return;

            const roleIndex:{[key:string]:number} = {};
            dv.table.columns.forEach((column:any,index:number) => {
                const roles = (column && column.roles) || {};
                Object.keys(roles).forEach((role:string) => { if(roles[role]) roleIndex[role] = index; });
            });

            self.events.forEach((event:any,index:number) => {
                const row = dv.table.rows[event.rowIndex == null ? index : event.rowIndex] || [];
                const value = (role:string) => roleIndex[role] === undefined ? null : row[roleIndex[role]];

                // Regra funcional obrigatória: Data Específica > Data Quinzena > Sem Data.
                event.type = event.date ? "Data Exata" : (hasValue(event.q) ? "Quinzena" : "Sem Data");
                event.subfonte = hasValue(value("subfonte")) ? String(value("subfonte")) : "–";
                event.executor = value("executor");
                event.mcmv = value("mcmv");
                event.novoPac = value("novoPac");
                event.uh = value("uh");
                event.observacoes = value("observacoes");
            });

            if(self.state.expandedEventId === undefined) self.state.expandedEventId = null;
            this.render();
        } catch(error) {
            this.showRuntimeError("atualização dos dados", error);
        }
    }

    public render(): void {
        try {
            this.renderV3();
        } catch(error) {
            this.showRuntimeError("renderização", error);
        }
    }

    private renderV3(): void {
        const self:any = this;
        if(!self.root) return;
        self.root.classList.toggle("mcid-no-topbar", !self.settings.showTopBar);
        self.root.innerHTML = "";

        self.tooltip = document.createElement("div");
        self.tooltip.className = "mcid-tooltip";
        self.root.appendChild(self.tooltip);

        self.toast = document.createElement("div");
        self.toast.className = "mcid-toast";
        self.toast.textContent = "Evento selecionado para navegação";
        self.root.appendChild(self.toast);

        if(self.settings.showTopBar) {
            const top = document.createElement("div");
            top.className = "mcid-topbar";
            top.innerHTML = '<div class="mcid-brand"><span class="mcid-brand-logo-wrap"><img class="mcid-brand-logo" src="' + MCID_LOGO_DATA_URI + '" alt="Ministério das Cidades"></span></div><div class="mcid-top-meta">Calendário</div>';
            self.root.appendChild(top);
        } else {
            self.root.classList.add("mcid-no-topbar");
        }

        const shell = document.createElement("div");
        shell.className = "mcid-shell";
        self.root.appendChild(shell);

        const title = document.createElement("div");
        title.className = "mcid-titlebar";
        title.innerHTML = '<div class="mcid-titlewrap"><div class="mcid-title-icon"></div><div><div class="mcid-title">CALENDÁRIO DE EVENTOS</div><div class="mcid-subtitle">Visão mensal</div></div></div><div class="mcid-nav"><button class="mcid-btn prev" aria-label="Mês anterior">‹</button><div class="mcid-month-label" aria-label="Mês exibido">' + esc(MONTHS[self.state.month] + " de " + self.state.year) + '</div><button class="mcid-btn next" aria-label="Mês seguinte">›</button></div>';
        shell.appendChild(title);

        (title.querySelector(".prev") as HTMLElement).onclick = () => { self.state.category = null; self.state.expandedEventId = null; self.shiftMonth(-1); };
        (title.querySelector(".next") as HTMLElement).onclick = () => { self.state.category = null; self.state.expandedEventId = null; self.shiftMonth(1); };

        const grid = document.createElement("div");
        grid.className = "mcid-grid";
        shell.appendChild(grid);
        grid.appendChild(this.renderForecast());
        grid.appendChild(this.renderLegend());
        grid.appendChild(self.renderCalendar());
        grid.appendChild(this.renderEvents());

        if(self.settings.showFooter) {
            const footer = document.createElement("div");
            footer.className = "mcid-footer";
            footer.innerHTML = '<span>MINISTÉRIO DAS CIDADES</span><span>Calendário MCid • V3</span>';
            shell.appendChild(footer);
        }
    }

    public renderForecast(): HTMLElement {
        const self:any = this;
        const box = document.createElement("div");
        box.className = "mcid-card mcid-forecast";
        box.innerHTML = '<div class="mcid-section-head"><div class="mcid-section-title">PREVISÕES SEM DATA DEFINIDA</div><span class="mcid-section-rule"></span><div class="mcid-info">Eventos sem data específica aparecem por quinzena e não são posicionados em um dia exato.</div></div>';

        const cards = document.createElement("div");
        cards.className = "mcid-forecast-cards";
        [["q1","1ª Quinzena","▦",""],["q2","2ª Quinzena","▦","q2"],["nodate","Sem Data","—","nodate"]].forEach((item:string[]) => {
            const n = self.forecastEvents(item[0]).length;
            const button = document.createElement("button");
            button.className = "mcid-forecast-btn " + item[3] + (self.state.forecast === item[0] ? " active" : "");
            button.innerHTML = '<span class="mcid-ficon">' + item[2] + '</span><span><div class="mcid-fname">' + item[1] + '</div><div class="mcid-fcount">' + n + ' ' + (n === 1 ? "evento" : "eventos") + '</div></span>';
            button.onclick = () => {
                self.state.forecast = self.state.forecast === item[0] ? null : item[0];
                self.state.day = null;
                self.state.category = null;
                self.state.expandedEventId = null;
                this.render();
            };
            cards.appendChild(button);
        });
        box.appendChild(cards);
        return box;
    }

    public renderLegend(): HTMLElement {
        const self:any = this;
        const box = document.createElement("div");
        box.className = "mcid-card mcid-legend";
        box.innerHTML = '<div class="mcid-section-head mcid-legend-head"><div class="mcid-section-title">LEGENDA</div><span class="mcid-section-rule"></span></div>';
        const list = document.createElement("div");
        list.className = "mcid-legend-list";
        CAT_ORDER.forEach((category:string) => {
            const item = document.createElement("div");
            item.className = "mcid-legend-item" + (self.state.category === category && !self.state.day ? " active" : "");
            item.innerHTML = '<span class="mcid-swatch" style="background:' + COLORS[category] + '"></span><span>' + esc(LEGEND_LABELS[category] || category) + '</span>';
            item.onclick = () => {
                self.state.category = self.state.category === category ? null : category;
                self.state.forecast = null;
                self.state.expandedEventId = null;
                this.render();
            };
            list.appendChild(item);
        });
        box.appendChild(list);
        return box;
    }

    public renderCalendar(): HTMLElement {
        const self:any = this;
        const box = document.createElement("div");
        box.className = "mcid-card mcid-calendar";

        const weekHead = document.createElement("div");
        weekHead.className = "mcid-week-head";
        ["DOM","SEG","TER","QUA","QUI","SEX","SÁB"].forEach((weekday:string) => {
            const el = document.createElement("div");
            el.textContent = weekday;
            weekHead.appendChild(el);
        });
        box.appendChild(weekHead);

        const days = document.createElement("div");
        days.className = "mcid-days";
        box.appendChild(days);

        const first = new Date(self.state.year,self.state.month,1);
        const start = new Date(self.state.year,self.state.month,1-first.getDay());
        const exactEvents = self.events.filter((event:any) => event.type === "Data Exata" && event.date);

        const keyDate = (date:Date):string => date.getFullYear() + "-" + pad2(date.getMonth()+1) + "-" + pad2(date.getDate());

        for(let index=0; index<42; index++) {
            const date = new Date(start.getFullYear(),start.getMonth(),start.getDate()+index);
            const events = exactEvents.filter((event:any) => keyDate(event.date) === keyDate(date));
            const cell = document.createElement("div");
            cell.className = "mcid-day" + (date.getMonth() !== self.state.month ? " out" : "") + (self.state.day && sameDate(self.state.day,date) ? " selected" : "");
            cell.innerHTML = '<div class="mcid-day-num">' + date.getDate() + '</div>';

            if(events.length > 0) {
                const categories = CAT_ORDER.filter((category:string) => events.some((event:any) => event.cat === category));
                const isSelectedDay = !!(self.state.day && sameDate(self.state.day,date));

                // Regra visual vigente: células usam somente os quadradinhos por categoria.
                // A quantidade permanece no tooltip; o texto total aparece somente no próprio dia
                // selecionado, para qualquer quantidade de eventos (1 ou mais), em todas as
                // seleções que mantenham esse dia ativo, inclusive clique em indicador/categoria.
                if(isSelectedDay) {
                    const total = document.createElement("div");
                    total.className = "mcid-total-label";
                    total.textContent = events.length + " " + (events.length === 1 ? "evento" : "eventos");
                    cell.appendChild(total);
                }

                const indicators = document.createElement("div");
                indicators.className = "mcid-category-indicators";
                categories.forEach((category:string) => {
                    const categoryEvents = events.filter((event:any) => event.cat === category);
                    const indicator = document.createElement("button");
                    indicator.className = "mcid-category-indicator";
                    indicator.setAttribute("aria-label", category + " — " + categoryEvents.length + " " + (categoryEvents.length === 1 ? "evento" : "eventos"));
                    indicator.title = category + " — " + categoryEvents.length + " " + (categoryEvents.length === 1 ? "evento" : "eventos");
                    indicator.innerHTML = '<span class="mcid-category-square" style="background:' + COLORS[category] + '"></span>';
                    indicator.onmouseenter = (evt:any) => self.showTip(evt,category,categoryEvents);
                    indicator.onmousemove = (evt:any) => self.moveTip(evt);
                    indicator.onmouseleave = () => self.hideTip();
                    indicator.onclick = (evt:any) => {
                        evt.stopPropagation();
                        self.state.day = date;
                        self.state.category = category;
                        self.state.forecast = null;
                        self.state.expandedEventId = null;
                        this.render();
                    };
                    indicators.appendChild(indicator);
                });
                cell.appendChild(indicators);
            }

            cell.onclick = () => {
                self.state.day = self.state.day && sameDate(self.state.day,date) ? null : date;
                self.state.forecast = null;
                self.state.expandedEventId = null;
                if(!self.state.day) self.state.category = null;
                this.render();
            };
            days.appendChild(cell);
        }
        return box;
    }

    private isPastMonth(): boolean {
        const self:any = this;
        const today = new Date();
        return new Date(self.state.year,self.state.month,1) < new Date(today.getFullYear(),today.getMonth(),1);
    }

    public contextTitle(): string {
        const self:any = this;
        if(self.state.forecast === "q1") return (self.state.category ? self.state.category.toUpperCase() + " • " : "") + "1ª QUINZENA • " + MONTHS[self.state.month].toUpperCase() + "/" + self.state.year;
        if(self.state.forecast === "q2") return (self.state.category ? self.state.category.toUpperCase() + " • " : "") + "2ª QUINZENA • " + MONTHS[self.state.month].toUpperCase() + "/" + self.state.year;
        if(self.state.forecast === "nodate") return (self.state.category ? self.state.category.toUpperCase() + " • " : "") + "SEM DATA";
        if(self.state.day && self.state.category) return self.state.category.toUpperCase() + " • " + fmtDate(self.state.day);
        if(self.state.day) return "EVENTOS DE " + fmtDate(self.state.day);
        if(self.state.category) return self.state.category.toUpperCase();
        return this.isPastMonth() ? "EVENTOS" : "PRÓXIMOS EVENTOS";
    }

    private eventKey(event:any): string {
        return hasValue(event.id) ? String(event.id) : "row-" + event.rowIndex;
    }

    private programLabel(event:any): string {
        const items:string[] = [];
        if(isYes(event.mcmv)) items.push("Minha Casa, Minha Vida");
        if(isYes(event.novoPac)) items.push("Novo PAC");
        return items.join(" • ");
    }

    private createEventDetails(event:any): HTMLElement {
        const details = document.createElement("div");
        details.className = "mcid-event-details";
        let html = '<div class="mcid-detail-line"><span class="mcid-detail-label">Secretaria:</span><span class="mcid-detail-value">' + esc(event.secretaria || "–") + '</span></div>';
        html += '<div class="mcid-detail-line"><span class="mcid-detail-label">Fonte:</span><span class="mcid-detail-value">' + esc(event.subfonte || "–") + '</span></div>';
        if(hasValue(event.executor)) html += '<div class="mcid-detail-line"><span class="mcid-detail-label">Executor:</span><span class="mcid-detail-value">' + esc(event.executor) + '</span></div>';
        const program = this.programLabel(event);
        if(program) html += '<div class="mcid-detail-line"><span class="mcid-detail-label">Programa:</span><span class="mcid-detail-value">' + esc(program) + '</span></div>';
        if(hasValue(event.uh)) html += '<div class="mcid-detail-line"><span class="mcid-detail-label">UH:</span><span class="mcid-detail-value">' + esc(event.uh) + '</span></div>';
        if(hasValue(event.observacoes)) html += '<div class="mcid-detail-line mcid-detail-observacoes"><span class="mcid-detail-label">Observações:</span><span class="mcid-detail-value">' + esc(event.observacoes) + '</span></div>';
        details.innerHTML = html;
        details.onclick = (evt:MouseEvent) => evt.stopPropagation();
        details.ondblclick = (evt:MouseEvent) => evt.stopPropagation();
        return details;
    }

    private applyExpandedState(list:HTMLElement,row:HTMLElement,body:HTMLElement,event:any): void {
        const self:any = this;
        const key = this.eventKey(event);
        const opening = self.state.expandedEventId !== key;
        const opened = list.querySelectorAll(".mcid-event-row.expanded");
        Array.prototype.forEach.call(opened,(other:HTMLElement) => {
            other.classList.remove("expanded");
            const old = other.querySelector(".mcid-event-details");
            if(old && old.parentNode) old.parentNode.removeChild(old);
        });
        self.state.expandedEventId = opening ? key : null;
        if(opening) {
            row.classList.add("expanded");
            body.appendChild(this.createEventDetails(event));
            try { row.scrollIntoView({block:"nearest",behavior:"smooth"}); } catch(_e) {}
        }
    }

    public renderEvents(): HTMLElement {
        const self:any = this;
        const events:any[] = self.contextEvents();
        const box = document.createElement("div");
        box.className = "mcid-card mcid-events";

        const validExpanded = events.some((event:any) => this.eventKey(event) === self.state.expandedEventId);
        if(!validExpanded) self.state.expandedEventId = null;

        const head = document.createElement("div");
        head.className = "mcid-events-head";
        head.innerHTML = '<div class="mcid-events-heading"><div class="mcid-events-title">' + esc(this.contextTitle()) + '</div><div class="mcid-events-sub">' + events.length + ' ' + (events.length === 1 ? "evento" : "eventos") + '</div></div>';
        box.appendChild(head);

        const summary = document.createElement("div");
        summary.className = "mcid-summary";
        const showSummary = !!(self.state.day || self.state.forecast);
        let summaryBase:any[] = events;
        if(self.state.forecast) summaryBase = self.forecastEvents(self.state.forecast);
        else if(self.state.day) summaryBase = self.events.filter((event:any) => event.type === "Data Exata" && sameDate(event.date,self.state.day));

        if(showSummary) {
            CAT_ORDER.forEach((category:string) => {
                const n = summaryBase.filter((event:any) => event.cat === category).length;
                if(!n) return;
                const button = document.createElement("button");
                button.className = "mcid-summary-btn" + (self.state.category === category ? " active" : "");
                button.innerHTML = '<div class="mcid-summary-name" style="color:' + COLORS[category] + '">' + category + '</div><div class="mcid-summary-count">' + n + '</div>';
                button.onclick = () => {
                    self.state.category = self.state.category === category ? null : category;
                    self.state.expandedEventId = null;
                    this.render();
                };
                summary.appendChild(button);
            });
        }
        box.appendChild(summary);

        const list = document.createElement("div");
        list.className = "mcid-event-list";
        if(!events.length) {
            list.innerHTML = '<div class="mcid-empty">Nenhum evento neste contexto.</div>';
        } else {
            events.forEach((event:any) => {
                const row = document.createElement("div");
                const key = this.eventKey(event);
                row.className = "mcid-event-row" + (self.state.expandedEventId === key ? " expanded" : "");

                const chip = event.date
                    ? ('<strong>' + pad2(event.date.getDate()) + '</strong>' + MONTHS[event.date.getMonth()].slice(0,3).toUpperCase())
                    : ('<strong>—</strong>' + (event.type === "Quinzena" ? (event.qn + "ª Q") : "SEM"));

                const dateCell = document.createElement("div");
                dateCell.className = "mcid-date-chip";
                dateCell.innerHTML = chip;

                const body = document.createElement("div");
                body.className = "mcid-event-body";
                body.innerHTML = '<div class="mcid-event-line1"><span class="mcid-event-category-bar" style="background:' + COLORS[event.cat] + '"></span><span class="mcid-event-category-text">' + esc(event.tip) + '</span></div><div class="mcid-event-line2">' + esc(event.project) + '</div><div class="mcid-event-line3">' + esc(event.city) + ' / ' + esc(event.uf) + '</div>';

                if(self.state.expandedEventId === key) body.appendChild(this.createEventDetails(event));

                row.appendChild(dateCell);
                row.appendChild(body);
                row.title = "1 clique: expandir detalhes • duplo clique: abrir Agenda de Eventos";
                row.onclick = (evt:MouseEvent) => {
                    if(evt.detail && evt.detail > 1) return;
                    self.selectEvent(event,false);
                    this.applyExpandedState(list,row,body,event);
                };
                row.ondblclick = (evt:MouseEvent) => {
                    evt.preventDefault();
                    evt.stopPropagation();
                    self.selectEvent(event,true);
                };
                list.appendChild(row);
            });
        }
        box.appendChild(list);
        return box;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return super.getFormattingModel();
    }
}
