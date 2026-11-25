"""Mitchell Lab browser-facing analysis tools (COSINOR intentionally disabled)."""
from __future__ import annotations

import hashlib
import hmac
import os
from pathlib import Path

import pandas as pd
import streamlit as st

from analysis_core import (AnalysisError, convert_io, manhattan_run, prepare_venn,
                           qba_protocol_labels, qba_run, venn_run, validate_stub)

st.set_page_config(page_title='Mitchell Lab | Analysis Portal', page_icon='🧪',layout='wide')


def gate():
    secret=os.environ.get('APP_PASSWORD','')
    if os.environ.get('APP_REQUIRE_PASSWORD')=='1' and not secret:
        st.error('App is configured to require a password, but APP_PASSWORD is missing.')
        st.stop()
    if not secret:
        return
    if st.session_state.get('access_granted'):
        return
    st.title('Mitchell Lab Analysis Portal')
    with st.form('access'):
        supplied=st.text_input('Access password',type='password')
        if st.form_submit_button('Unlock'):
            if hmac.compare_digest(supplied,secret):
                st.session_state['access_granted']=True
                st.rerun()
            else:
                st.error('Incorrect password.')
    st.stop()


def decode_uploads(uploads):
    return [(x.name,x.getvalue()) for x in uploads]


def sheet_names(raw):
    try:
        return pd.ExcelFile(raw).sheet_names
    except Exception:
        return []


def download_zip(data, filename):
    st.download_button('Download results ZIP',data=data,file_name=filename,
                       mime='application/zip',use_container_width=True)


def execute(func,*,key):
    st.session_state.pop(key, None)
    try:
        with st.spinner('Analyzing uploaded files...'):
            result=func()
        st.session_state[key]=result
        st.success('Analysis completed; download the results below.')
    except AnalysisError as exc:
        st.error(str(exc))
    except Exception as exc:
        st.error(f'Analysis stopped: {type(exc).__name__}: {exc}')


gate()
st.title('Mitchell Lab | Analysis Portal')
st.caption('Browser interface for the existing lab pipelines · COSINOR is under development')
st.info('Use de-identified research files only. This starter is not an institutional data-compliance or authentication system.')

tab_convert,tab_venn,tab_manhattan,tab_qba,tab_cosinor = st.tabs([
    'Convert Txt', 'Venn diagrams', 'Manhattan plots', 'Quiet Breathing Analysis', 'COSINOR (paused)'])

with tab_convert:
    st.subheader('IOX text → cleaned Excel')
    st.write('Adapted from the supplied `edu_iox_dataclean_notebook.ipynb`. The output ZIP includes raw Excel conversions and cleaned workbooks wherever the original cleaning schema is recognized.')
    files=st.file_uploader('IOX/TXT files',type=['txt','iox'],accept_multiple_files=True,key='convert_upload')
    if st.button('Convert files',disabled=not files,key='do_convert'):
        execute(lambda: convert_io(decode_uploads(files)),key='convert_result')
    if 'convert_result' in st.session_state:
        result,log=st.session_state['convert_result']
        st.code('\n'.join(log))
        download_zip(result,'mitchell_converted_io.zip')

with tab_venn:
    st.subheader('Venn set comparisons')
    st.write('Uses the supplied `venn_functions.R` and `analysis_runner.R`. One set per selected worksheet; a one-column headerless workbook is supported by the original reader.')
    files=st.file_uploader('Excel workbook(s)',type=['xlsx','xls'],accept_multiple_files=True,key='venn_upload')
    c1,c2,c3=st.columns(3)
    with c1: prefix=st.text_input('Output prefix',value='Venn',key='venn_prefix')
    with c2: id_header=st.text_input('CpG/ID column (optional)',key='venn_id')
    with c3: gene_header=st.text_input('Gene Name column (optional)',key='venn_gene')
    sheets=st.text_input('Sheet names separated by semicolons (blank = all)',key='venn_sheets')
    requested=[s.strip() for s in sheets.split(';') if s.strip()]
    signature=hashlib.sha256(repr(([(f.name,len(f.getvalue()),hashlib.sha256(f.getvalue()).hexdigest()) for f in (files or [])], prefix,id_header,gene_header,requested)).encode()).hexdigest()
    if st.button('1. Inspect sets',disabled=not files,key='inspect_venn'):
        execute(lambda: prepare_venn(decode_uploads(files),prefix,requested,id_header,gene_header),key='venn_preview_result')
        st.session_state['venn_signature']=signature
    if 'venn_preview_result' in st.session_state and st.session_state.get('venn_signature')==signature:
        preview=st.session_state['venn_preview_result']
        st.dataframe(pd.DataFrame(preview),hide_index=True)
        options=[str(r['name']) for r in preview]
        chosen=st.multiselect('Select 2–4 sets',options,default=options[:min(2,len(options))],max_selections=4)
        all_pairs=st.checkbox('Generate all pairwise comparisons among selected sets',value=False)
        if st.button('2. Generate diagrams',disabled=len(chosen)<2,key='run_venn'):
            execute(lambda: venn_run(decode_uploads(files),chosen,prefix=prefix,sheets=requested,
                                     id_column=id_header,gene_column=gene_header,all_pairs=all_pairs),key='venn_result')
    elif 'venn_preview_result' in st.session_state:
        st.warning('Inputs or settings changed; inspect sets again.')
    if 'venn_result' in st.session_state: download_zip(st.session_state['venn_result'],'mitchell_venn_results.zip')

with tab_manhattan:
    st.subheader('Manhattan plots')
    st.write('Uses the uploaded `manhattan_functions.R` and `analysis_runner.R`. Input p-values should already be adjusted; this tool does not calculate FDR.')
    files=st.file_uploader('Excel workbook(s)',type=['xlsx','xls'],accept_multiple_files=True,key='manhattan_upload')
    ca,cb=st.columns(2)
    with ca:
        prefix=st.text_input('Output prefix',value='Manhattan',key='man_prefix')
        chromosome=st.text_input('Chromosome column',value='chr')
        position=st.text_input('Position column',value='pos')
        qval=st.text_input('Adjusted p-value column',value='adj.P.Val')
    with cb:
        cpg=st.text_input('CpG / point label column',value='CpG')
        autosomes=st.number_input('Autosome count',min_value=1,max_value=100,value=22,step=1)
        cutoff=st.number_input('Adjusted p threshold',min_value=0.000001,max_value=1.0,value=0.05,format='%.6f')
        top=st.number_input('Labels for lowest p-values',min_value=0,max_value=250,value=10)
    sheets=st.text_input('Sheet names separated by semicolons (blank = all)',key='man_sheets')
    if st.button('Generate Manhattan plots',disabled=not files,key='run_manhattan'):
        execute(lambda: manhattan_run(decode_uploads(files),prefix=prefix,
                        columns=[chromosome,position,qval,cpg],autosomes=autosomes,
                        sheets=[s.strip() for s in sheets.split(';') if s.strip()],cutoff=cutoff,
                        top_labels=top),key='man_result')
    if 'man_result' in st.session_state: download_zip(st.session_state['man_result'],'mitchell_manhattan_results.zip')

with tab_qba:
    st.subheader('Quiet Breathing Analysis (QBA)')
    st.write('Run the supplied `QB_Breathing_Tag_Updated_Fin.Rmd` on cleaned QB Excel workbooks. Breath tagging, classification precedence, thresholds, and the final two minutes of challenge blocks are retained from that source.')
    st.caption('Upload one to five cleaned .xlsx workbooks. All are processed independently using the same protocol classification. Outputs include QB-filter workbooks, average tables, and breathing-classification summaries.')
    files=st.file_uploader('Cleaned Quiet Breathing Excel workbook(s)',type=['xlsx'],accept_multiple_files=True,key='qba_files')
    tag=st.text_input('Rat light type / output prefix',value='LD',key='qba_tag')
    labels=[]
    qba_error=None
    if files:
        try:
            labels=qba_protocol_labels(decode_uploads(files))
        except AnalysisError as exc:
            qba_error=str(exc)
    if qba_error:
        st.error(qba_error)
    if labels:
        st.markdown('**Classify every Protocol Type label:**')
        st.caption('Each distinct label must belong to exactly one category. Challenge blocks use the original Rmd rule: analyze their last two minutes. Excluded labels are not treated as baseline or challenge.')
        # Changing the uploaded workbooks resets the defaults; edits persist on normal reruns.
        qba_signature=hashlib.sha256(repr([
            (f.name, hashlib.sha256(f.getvalue()).hexdigest()) for f in files
        ]).encode()).hexdigest()[:16]
        known_base={'Quiet Breathing','Normoxia Baseline'}
        default_baseline=[x for x in labels if x in known_base]
        default_excluded=[x for x in labels if x == 'Acclimation']
        default_challenge=[x for x in labels if x not in default_baseline and x not in default_excluded]
        base_col, challenge_col, exclude_col = st.columns(3)
        with base_col:
            baseline=st.multiselect('Baseline', labels, default=default_baseline,
                                    key='qba_baseline_'+qba_signature)
        with challenge_col:
            challenge=st.multiselect('Challenge', labels, default=default_challenge,
                                     key='qba_challenge_'+qba_signature)
        with exclude_col:
            excluded=st.multiselect('Excluded', labels, default=default_excluded,
                                    key='qba_excluded_'+qba_signature)
        classified=baseline+challenge+excluded
        duplicate=sorted(set(x for x in classified if classified.count(x)>1))
        missing=[label for label in labels if label not in classified]
        if duplicate:
            st.error('Labels in multiple categories: '+', '.join(map(repr, duplicate)))
        if missing:
            st.error('Unassigned protocol labels: '+', '.join(map(repr,missing)))
        if not baseline:
            st.warning('Select at least one Baseline protocol before running QBA.')
        ready=bool(baseline) and not duplicate and not missing and not qba_error
        result_signature=hashlib.sha256(repr((qba_signature,tag,baseline,challenge,excluded)).encode()).hexdigest()
        if (st.session_state.get('qba_result_signature') != result_signature):
            st.session_state.pop('qba_result',None)
        if st.button('Run Quiet Breathing Analysis',disabled=not ready,key='qba_run_button'):
            execute(lambda: qba_run(decode_uploads(files),prefix=tag,
                                    baseline=baseline, challenge=challenge,
                                    excluded=excluded),key='qba_result')
            if 'qba_result' in st.session_state:
                st.session_state['qba_result_signature']=result_signature
        if 'qba_result' in st.session_state:
            download_zip(st.session_state['qba_result'],'mitchell_quiet_breathing_results.zip')
    else:
        st.session_state.pop('qba_result',None)
    st.caption('Research validation: compare representative outputs with the original standalone Rmd before using the results in a publication or clinical decision. The live Seagate desktop patches were not uploaded into this package.')

with tab_cosinor:
    st.subheader('COSINOR Modeling')
    st.warning('Temporarily disabled in the web portal. The four templates and helper scripts are preserved in the master source archive for ongoing development.')
    st.write('No analysis is executed from this tab. Re-enable only after consistent worksheet/range, group, ZT24, outlier, and unit handling is validated.')

st.divider()
st.caption('Files are processed in isolated temporary folders; downloads stay in each browser session until the session ends. Avoid uploading protected or identifiable data to an unapproved host.')
