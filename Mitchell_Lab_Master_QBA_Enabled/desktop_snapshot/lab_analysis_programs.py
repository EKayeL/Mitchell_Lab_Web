"""Venn and Manhattan app screens; run R in a subprocess, UI only on Tk thread."""
import csv
import json
import math
import os
import queue
import re
import subprocess
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

ASSETS = Path(__file__).resolve().parent / 'analysis_assets'
COMMON = [
    ('prefix','Output filename prefix','Analysis'),
    ('sheets','Sheets (semicolon-separated; blank = all)',''),
    ('dpi','Image resolution (DPI)','300'),
]
VENN = [
    ('id_column','ID/CpG column (blank = automatic)',''),
    ('gene_column','Gene Name column (blank = automatic/optional)',''),
    ('excel_folder','Excel output subfolder','Sets and Intersections'),
    ('plot_folder','Plot output subfolder','Venn Diagrams'),
    ('colors','Set colors (four comma-separated hex colors)','#999999,#E69F00,#56B4E9,#009E73'),
    ('count_size','Region-count text size (cex)','3'),
    ('label_size','Set-label text size (cex)','1'),
    ('min_width','Minimum plot width (inches)','8'),
    ('max_width','Maximum plot width (inches)','16'),
    ('min_height','Minimum plot height (inches)','8'),
    ('max_height','Maximum plot height (inches)','14'),
]
MANHATTAN = [
    ('chr_col','Chromosome column','chr'),('pos_col','Position column','pos'),
    ('p_col','Adjusted p-value column','adj.P.Val'),('id_col','CpG/label column','CpG'),
    ('title','Plot title','DNA Methylation Manhattan Plot'),
    ('suffix','Output filename suffix','DNA_methylation_manhattan'),
    ('x_label','X-axis label','Chromosome'),('y_label','Y-axis label','-log10(Adjusted P-value)'),
    ('cutoff','Adjusted p-value threshold','0.05'),('top_labels','Top hits to label (0 = none)','10'),
    ('autosomes','Number of autosomes (human 22; rat 20)','22'),
    ('colors','Alternating colors (two comma-separated hex colors)','#4D4D4D,#87CEEB'),
    ('threshold_color','Threshold-line color (hex)','#FF0000'),
    ('alpha','Point opacity (0–1)','0.75'),('point_size','Point size','0.8'),
    ('label_size','Hit-label size','3'),('font_size','Plot base font size','14'),
    ('chr_width','Chromosome display width','1'),('chr_gap','Gap between chromosomes','0.12'),
    ('min_width','Minimum plot width (inches)','10'),('max_width','Maximum plot width (inches)','24'),
    ('min_height','Minimum plot height (inches)','5'),('max_height','Maximum plot height (inches)','12'),
]

def r_literal(value):
    if value is None: return 'NULL'
    if isinstance(value, bool): return 'TRUE' if value else 'FALSE'
    if isinstance(value, (int,float)):
        if not math.isfinite(value): raise ValueError('Nonfinite configuration value')
        return str(value)
    if isinstance(value, str): return json.dumps(value, ensure_ascii=False)
    if isinstance(value, (list,tuple)): return 'list('+','.join(r_literal(v) for v in value)+')'
    if isinstance(value, dict): return 'list('+','.join(r_literal(k)+'='+r_literal(v) for k,v in value.items())+')'
    raise TypeError(type(value))

def filename_component(value, label):
    value=value.strip()
    if not value or value in ('.','..') or re.search(r'[\x00-\x1f/\\:*?"<>|]',value):
        raise ValueError(label+' must be a nonempty name without slashes or filename control characters.')
    if len(value)>90: raise ValueError(label+' must have at most 90 characters.')
    return value

def validate_options(kind, raw):
    cfg=dict(raw)
    cfg['prefix']=filename_component(cfg['prefix'],'Output prefix')
    cfg['sheets']=[s.strip() for s in cfg['sheets'].split(';') if s.strip()]
    if len(set(cfg['sheets']))!=len(cfg['sheets']): raise ValueError('Sheet names must not repeat.')
    bounds={'dpi':(72,1200,True),'min_width':(1,40,False),'max_width':(1,40,False),
            'min_height':(1,40,False),'max_height':(1,40,False),'label_size':(.1,20,False)}
    if kind=='venn': bounds['count_size']=(.1,20,False)
    else: bounds.update({'cutoff':(1e-300,1,False),'top_labels':(0,500,True),'autosomes':(1,100,True),
        'alpha':(0,1,False),'point_size':(.01,20,False),'font_size':(6,48,False),
        'chr_width':(.1,10,False),'chr_gap':(0,10,False)})
    for key,(low,high,integer) in bounds.items():
        try: value=float(cfg[key])
        except ValueError: raise ValueError(key+' must be numeric.')
        if not math.isfinite(value) or not low<=value<=high or (integer and not value.is_integer()):
            raise ValueError(f'{key} must be {"a whole number" if integer else "a number"} between {low} and {high}.')
        cfg[key]=int(value) if integer else value
    for axis in ('width','height'):
        if cfg['min_'+axis]>cfg['max_'+axis]: raise ValueError('Minimum '+axis+' cannot exceed maximum.')
    if cfg['max_width']*cfg['max_height']*cfg['dpi']**2>150_000_000:
        raise ValueError('Maximum image dimensions exceed 150 million pixels. Reduce DPI or dimensions.')
    cfg['colors']=[s.strip() for s in cfg['colors'].split(',')]
    count=4 if kind=='venn' else 2
    if len(cfg['colors'])!=count or any(not re.fullmatch(r'#[0-9a-fA-F]{6}',s) for s in cfg['colors']):
        raise ValueError(f'Provide exactly {count} colors in #RRGGBB format, separated by commas.')
    if kind=='venn':
        for key in ('excel_folder','plot_folder'): cfg[key]=filename_component(cfg[key],key)
        if cfg['excel_folder'].casefold()==cfg['plot_folder'].casefold(): raise ValueError('Use distinct plot and Excel subfolder names.')
    else:
        cfg['columns']=[cfg.pop(k).strip() for k in ('chr_col','pos_col','p_col','id_col')]
        if len(set(cfg['columns']))!=4 or any(not v for v in cfg['columns']): raise ValueError('Provide four distinct, nonempty column names.')
        cfg['suffix']=filename_component(cfg['suffix'],'Filename suffix')
        if not re.fullmatch(r'#[0-9a-fA-F]{6}',cfg['threshold_color']): raise ValueError('Threshold color must be #RRGGBB.')
    return cfg


def make_analysis_frames(AnalysisFrame, AppButton, find_rscript, DND_FILES):
    class ProgramFrame(AnalysisFrame):
        def __init__(self,parent,app):
            super().__init__(parent,bg=app.bg_color)
            self.app=app; self.selected_files=[]; self._busy=False
            self._events=queue.Queue(); self._process=None; self._cancel=threading.Event()
            self._job=None; self.entries={}; self._popup=None
            self.center=tk.Frame(self,bg=app.bg_color)
            self.center.pack(fill='both',expand=True); app.register_bg_widget(self.center)
            viewport=tk.Frame(self.center,bg=app.bg_color); viewport.pack(fill='both',expand=True)
            app.register_bg_widget(viewport)
            canvas=tk.Canvas(viewport,bg=app.bg_color,highlightthickness=0)
            canvas.pack(side='left',fill='both',expand=True); app.register_bg_widget(canvas)
            scrollbar=ttk.Scrollbar(viewport,orient='vertical',command=canvas.yview); scrollbar.pack(side='right',fill='y')
            canvas.configure(yscrollcommand=scrollbar.set)
            form=tk.Frame(canvas,bg=app.bg_color); app.register_bg_widget(form)
            window=canvas.create_window(0,0,window=form,anchor='nw')
            canvas.bind('<Configure>',lambda e: canvas.itemconfigure(window,width=e.width))
            form.bind('<Configure>',lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
            self._label(form, 'Venn Diagram Analysis' if self.KIND=='venn' else 'Manhattan Plot Analysis',True).pack(pady=8)
            self._label(form,'Upload Excel workbooks (.xlsx, .xls). Use the scrollbar to see all options.').pack()
            self.file_list=tk.Listbox(form,height=5,exportselection=False,font=app.small_btn_font)
            self.file_list.pack(fill='x',padx=20,pady=8)
            if DND_FILES and hasattr(self.file_list,'drop_target_register'):
                self.file_list.drop_target_register(DND_FILES)
                self.file_list.dnd_bind('<<Drop>>',lambda e:self._add_files(self.tk.splitlist(e.data)))
            row=tk.Frame(form,bg=app.bg_color); row.pack(); app.register_bg_widget(row)
            self._button(row,'Browse Files',self._browse).pack(side='left',padx=6)
            self._button(row,'Clear',self._clear).pack(side='left',padx=6)
            options=COMMON+(VENN if self.KIND=='venn' else MANHATTAN)
            fields=tk.Frame(form,bg=app.bg_color); fields.pack(fill='x',padx=20,pady=12); app.register_bg_widget(fields)
            fields.columnconfigure(1,weight=1)
            for i,(key,label,default) in enumerate(options):
                if key=='prefix': default='Venn' if self.KIND=='venn' else 'Manhattan'
                self._label(fields,label).grid(row=i,column=0,sticky='w',padx=(0,12),pady=4)
                variable=tk.StringVar(value=default); self.entries[key]=variable
                entry=tk.Entry(fields,textvariable=variable,font=app.small_btn_font)
                entry.grid(row=i,column=1,sticky='ew',pady=4)
            if self.KIND=='manhattan':
                self._label(form,'Uses adjusted p-values and equal-width chromosome positions, as in your R script.').pack(pady=6)
            else:
                self._label(form,'Run Files loads sets, then opens a popup to name and queue 2–4-set comparisons.').pack(pady=6)
            self._button(form,'Install R Packages',self._install_packages).pack(pady=10)
            bottom=tk.Frame(self,bg=app.bg_color); bottom.pack(side='bottom',fill='x'); app.register_bg_widget(bottom)
            self.status_var=tk.StringVar(value='Ready.')
            status=self._label(bottom,''); status.configure(textvariable=self.status_var,wraplength=760); status.pack(pady=5)
            self.progress=ttk.Progressbar(bottom,mode='indeterminate'); self.progress.pack(fill='x',padx=20)
            bar=tk.Frame(bottom,bg=app.bg_color); bar.pack(pady=8); app.register_bg_widget(bar)
            self._button(bar,'Run Files',self.run_files).pack(side='left',padx=6)
            self._button(bar,'Stop',self._stop).pack(side='left',padx=6)
            self._button(bar,'Back',lambda:app.show_frame('run_select')).pack(side='left',padx=6)
            self.after(100,self._poll)

        def _label(self,parent,text,bold=False):
            label=tk.Label(parent,text=text,bg=self.app.bg_color,fg='white',
                font=self.app.btn_font if bold else self.app.small_btn_font,anchor='w',justify='left')
            self.app.register_bg_widget(label); return label

        def _button(self,parent,text,command):
            return AppButton(parent,text=text,font=self.app.small_btn_font,command=command,padx=12,pady=8)

        def on_show(self): pass

        def _browse(self):
            if self._busy: return
            self._add_files(filedialog.askopenfilenames(parent=self,title='Choose Excel workbooks'))

        def _add_files(self,paths):
            if self._busy: return
            bad=[]
            for value in paths:
                path=Path(value).expanduser().resolve()
                if path.suffix.lower() not in ('.xls','.xlsx') or not path.is_file(): bad.append(path.name); continue
                if str(path) not in self.selected_files: self.selected_files.append(str(path))
            self.file_list.delete(0,'end')
            for path in self.selected_files: self.file_list.insert('end',path)
            if bad: messagebox.showwarning('Files not added','Only existing .xls/.xlsx workbooks are accepted:\n'+'\n'.join(bad),parent=self)

        def _clear(self):
            if self._busy: return
            self.selected_files=[]; self.file_list.delete(0,'end'); self.status_var.set('Ready.')

        def _begin(self):
            self._busy=True; self._cancel.clear(); self.progress.start(12)

        def _poll(self):
            try:
                while True:
                    kind,data=self._events.get_nowait()
                    if kind=='status': self.status_var.set(data[-400:])
                    elif kind=='prepared':
                        self.progress.stop(); self.status_var.set('Select comparisons in the popup.'); self._choose_comparisons(data)
                    elif kind in ('done','error','cancelled'):
                        self._busy=False; self.progress.stop(); self.status_var.set(data)
                        if kind=='done': messagebox.showinfo('Complete',data,parent=self)
                        elif kind=='error': messagebox.showerror('Analysis error',data,parent=self)
            except queue.Empty: pass
            self.after(100,self._poll)

        def _stop(self):
            if not self._busy: return
            if self._popup and self._popup.winfo_exists():
                self._popup.destroy(); self._popup=None; self._busy=False
                self.status_var.set('Comparison selection cancelled.'); return
            self._cancel.set()
            proc=self._process
            if proc and proc.poll() is None: proc.terminate()
            self.status_var.set('Stopping…')

        def _launch(self,command,cwd,log):
            if self._cancel.is_set(): raise InterruptedError('Cancelled.')
            tail=[]
            with open(log,'a',encoding='utf-8') as handle:
                proc=subprocess.Popen(command,cwd=str(cwd),stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding='utf-8',errors='replace')
                self._process=proc
                if self._cancel.is_set(): proc.terminate()
                for line in proc.stdout:
                    handle.write(line); handle.flush(); tail.append(line); tail=tail[-30:]
                    if line.strip(): self._events.put(('status',line.strip()))
                code=proc.wait(); self._process=None
            if self._cancel.is_set(): raise InterruptedError('Cancelled. Any partial outputs are retained in '+str(cwd))
            if code: raise RuntimeError(''.join(tail)[-4500:]+'\n\nFull log: '+str(log))

        def _execute(self,cfg,prepared=False):
            try:
                config=Path(cfg['job'])/('prepare_config.R' if prepared else 'analysis_config.R')
                config.write_text('cfg <- '+r_literal(cfg)+'\n',encoding='utf-8')
                (Path(cfg['job'])/('prepare_options.json' if prepared else 'analysis_options.json')).write_text(json.dumps(cfg,indent=2,ensure_ascii=False),encoding='utf-8')
                self._launch([find_rscript(),'--vanilla',str(ASSETS/'analysis_runner.R'),str(config)],Path(cfg['job']),Path(cfg['job'])/'run.log')
                if prepared:
                    with open(Path(cfg['job'])/'sets.tsv',encoding='utf-8-sig',newline='') as f: rows=list(csv.DictReader(f,delimiter='\t'))
                    self._events.put(('prepared',(cfg,rows)))
                else: self._events.put(('done','Analysis completed. Results saved in:\n'+cfg['job']))
            except InterruptedError as exc: self._events.put(('cancelled',str(exc)))
            except Exception as exc: self._events.put(('error',str(exc)))

        def run_files(self):
            if self._busy: return
            try:
                if not self.selected_files: raise ValueError('Upload at least one Excel workbook.')
                if any(not Path(p).is_file() for p in self.selected_files): raise ValueError('An uploaded file has moved or is unavailable.')
                cfg=validate_options(self.KIND,{k:v.get().strip() for k,v in self.entries.items()})
                find_rscript()
                if not (ASSETS/'analysis_runner.R').is_file(): raise ValueError('Analysis assets are missing. Re-run the installer.')
                default='App Venn Outputs' if self.KIND=='venn' else 'App Manhattan Outputs'
                root=self.resolve_output_folder(self.selected_files[0],default,self.output_folder)
                job=root/(cfg['prefix']+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
                job.mkdir(parents=True,exist_ok=False)
                cfg.update(kind=self.KIND,files=list(self.selected_files),job=str(job.resolve()),assets=str(ASSETS),stage='prepare' if self.KIND=='venn' else 'run')
                self._job=job; self._begin(); self.status_var.set('Loading workbooks…')
                threading.Thread(target=self._execute,args=(cfg,self.KIND=='venn'),daemon=True).start()
            except Exception as exc: messagebox.showerror('Cannot start',str(exc),parent=self)

        def _install_packages(self):
            if self._busy: return
            try: executable=find_rscript()
            except Exception as exc: messagebox.showerror('R required',str(exc),parent=self); return
            packages=['readxl','dplyr','VennDiagram','openxlsx','tibble','ggplot2','stringr','ggrepel']
            code='p <- unlist('+r_literal(packages)+'); u <- path.expand(Sys.getenv("R_LIBS_USER")); dir.create(u,recursive=TRUE,showWarnings=FALSE); .libPaths(c(u,.libPaths())); m <- p[!vapply(p,requireNamespace,logical(1),quietly=TRUE)]; if(length(m)) install.packages(m,lib=u,repos="https://cloud.r-project.org"); if(any(!vapply(p,requireNamespace,logical(1),quietly=TRUE))) stop("Some packages could not be installed"); cat("All analysis packages are available.\\n")'
            self._begin(); self.status_var.set('Checking/installing R packages…')
            def work():
                try:
                    logs=Path.home()/'Library'/'Logs'/'MitchellLabApp'; logs.mkdir(parents=True,exist_ok=True)
                    log=logs/('packages_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.log')
                    self._launch([executable,'--vanilla','-e',code],logs,log)
                    self._events.put(('done','R analysis packages are ready.'))
                except InterruptedError as exc: self._events.put(('cancelled',str(exc)))
                except Exception as exc: self._events.put(('error',str(exc)))
            threading.Thread(target=work,daemon=True).start()

        def _choose_comparisons(self,payload):
            cfg,rows=payload
            pop=tk.Toplevel(self); self._popup=pop; pop.title('Select Venn comparisons'); pop.bg_color=self.app.bg_color
            pop.configure(bg=self.app.bg_color); pop.geometry('1000x650'); pop.minsize(650,450); pop.transient(self.winfo_toplevel())
            aliases=[r['name'] for r in rows]; queued=[]
            note=tk.Label(pop,text='Select 2–4 sets, optionally rename them, then add comparisons to the queue.',font=self.app.small_btn_font)
            note.pack(pady=10)
            footer=tk.Frame(pop,bg=self.app.bg_color); footer.pack(side='bottom',fill='x',pady=10)
            body=tk.Frame(pop); body.pack(fill='both',expand=True,padx=12)
            style=ttk.Style(pop)
            style.configure('Analysis.Treeview',font=self.app.small_btn_font,rowheight=self.app.small_btn_font.metrics('linespace')+8)
            style.configure('Analysis.Treeview.Heading',font=self.app.small_btn_font)
            tree=ttk.Treeview(body,columns=('name','count','source'),show='headings',selectmode='extended',style='Analysis.Treeview')
            for col,title,width in [('name','Set label',220),('count','IDs',65),('source','Workbook / sheet',300)]:
                tree.heading(col,text=title); tree.column(col,width=width)
            scroll=ttk.Scrollbar(body,orient='vertical',command=tree.yview); tree.configure(yscrollcommand=scroll.set)
            tree.pack(side='left',fill='both',expand=True); scroll.pack(side='left',fill='y')
            for i,row in enumerate(rows): tree.insert('', 'end',iid=str(i),values=(aliases[i],row['count'],row['file']+' / '+row['sheet']))
            right=tk.Frame(body); right.pack(side='right',fill='both',expand=True,padx=8)
            tk.Label(right,text='Queued comparisons',font=self.app.small_btn_font).pack()
            queue_list=tk.Listbox(right,exportselection=False,font=self.app.small_btn_font); queue_list.pack(fill='both',expand=True)
            rename=tk.StringVar(); comparison_name=tk.StringVar()
            tk.Label(right,text='New label for one selected set').pack(); tk.Entry(right,textvariable=rename,font=self.app.small_btn_font).pack(fill='x')
            def rename_set():
                chosen=tree.selection()
                if len(chosen)!=1: messagebox.showinfo('Rename','Select exactly one set.',parent=pop); return
                name=rename.get().strip(); idx=int(chosen[0])
                if not name or len(name)>100 or '∩' in name or name.endswith(' Only') or any(ord(c)<32 for c in name) or any(name==a for i,a in enumerate(aliases) if i!=idx):
                    messagebox.showerror('Set label','Use a unique label (1–100 characters), without ∩, control characters, or a trailing " Only".',parent=pop); return
                aliases[idx]=name; tree.item(str(idx),values=(name,rows[idx]['count'],rows[idx]['file']+' / '+rows[idx]['sheet'])); refresh_queue()
            self._button(right,'Rename Set',rename_set).pack(pady=5)
            tk.Label(right,text='Comparison filename (blank = automatic)').pack(); tk.Entry(right,textvariable=comparison_name,font=self.app.small_btn_font).pack(fill='x')
            def refresh_queue():
                queue_list.delete(0,'end')
                for item in queued: queue_list.insert('end',item['name']+': '+' ∩ '.join(aliases[i-1] for i in item['indices']))
            def add():
                idx=sorted(int(v)+1 for v in tree.selection())
                if not 2<=len(idx)<=4: messagebox.showwarning('Comparison','Select between two and four sets (Command-click to select multiple).',parent=pop); return
                if any(q['indices']==idx for q in queued): messagebox.showinfo('Comparison','That comparison is already queued.',parent=pop); return
                try: name=filename_component(comparison_name.get().strip() or 'comparison_'+str(len(queued)+1),'Comparison filename')
                except ValueError as exc: messagebox.showerror('Name',str(exc),parent=pop); return
                queued.append({'indices':idx,'name':name}); refresh_queue(); tree.selection_remove(*tree.selection()); comparison_name.set('')
            def remove():
                selected=queue_list.curselection()
                if selected: queued.pop(selected[0]); refresh_queue()
            self._button(right,'Add Comparison',add).pack(pady=5)
            self._button(right,'Remove Selected',remove).pack(pady=5)
            def cancel():
                self._popup=None; pop.destroy(); self._busy=False; self.progress.stop(); self.status_var.set('Cancelled before analysis. No comparisons were run.')
            def run():
                if not queued: messagebox.showwarning('Comparisons','Add at least one comparison.',parent=pop); return
                cfg.update(stage='run',aliases=list(aliases),comparisons=list(queued))
                self._popup=None; pop.destroy(); self.progress.start(12)
                threading.Thread(target=self._execute,args=(cfg,False),daemon=True).start()
            self._button(footer,'Run Queued Comparisons',run).pack(side='left',padx=12)
            self._button(footer,'Cancel',cancel).pack(side='right',padx=12)
            pop.protocol('WM_DELETE_WINDOW',cancel); pop.grab_set()

    VennAnalysisFrame=type('VennAnalysisFrame',(ProgramFrame,),{'KIND':'venn'})
    ManhattanAnalysisFrame=type('ManhattanAnalysisFrame',(ProgramFrame,),{'KIND':'manhattan'})
    return VennAnalysisFrame,ManhattanAnalysisFrame
