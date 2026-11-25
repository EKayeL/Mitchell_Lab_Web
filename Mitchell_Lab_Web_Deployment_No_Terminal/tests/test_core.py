import unittest
from pathlib import Path
import sys
import tempfile
import zipfile
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analysis_core as core

class AdapterTests(unittest.TestCase):
    def test_r_literals_never_leave_string(self):
        val='foo";system("bad");#\\\\\n'
        encoded=core.r_quote(val)
        self.assertEqual(__import__('json').loads(encoded),val)

    def test_manhattan_configuration(self):
        with tempfile.TemporaryDirectory() as d:
            cfg=core.manhattan_config(Path(d),[Path(d)/'test.xlsx'],prefix='Manhattan',
                    columns=['chr','pos','adj.P.Val','CpG'],autosomes=20)
        self.assertEqual(cfg['autosomes'],20)
        self.assertEqual(cfg['kind'],'manhattan')
        self.assertEqual(cfg['columns'],['chr','pos','adj.P.Val','CpG'])

    def test_venn_configuration(self):
        with tempfile.TemporaryDirectory() as d:
            cfg=core.venn_config(Path(d),[Path(d)/'test.xlsx'],'Venn')
        self.assertEqual(cfg['stage'],'prepare')
        self.assertEqual(cfg['assets'],str(core.ASSETS))

    def test_qba_no_prompts(self):
        src=(core.ASSETS/'QB_Breathing_Tag_Updated_Fin.Rmd').read_text(encoding='utf8')
        code=core.make_qba_script(src,Path('/tmp/input.xlsx'),'LD',['Quiet Breathing'],['Acclimation'])
        self.assertNotIn('file.choose()',code)
        self.assertNotIn('readline(',code)
        self.assertIn('baseline_protocol <- unlist(',code)
        self.assertIn('is_postsigh',code)

    def test_txt_output_zip(self):
        payload, notes=core.convert_io([('foo.iox',b'A\tB\n1\t2\n')])
        with zipfile.ZipFile(__import__('io').BytesIO(payload)) as z:
            names=z.namelist()
            self.assertTrue(any(n.endswith('.xlsx') for n in names))
            self.assertIn('conversion_report.txt',names)
            self.assertTrue(any('raw conversion only' in n for n in notes))

    def test_qba_all_classifications_and_script(self):
        import io
        import pandas as pd
        sheet = io.BytesIO()
        pd.DataFrame({'Protocol Type':['Quiet Breathing','Challenge A','Acclimation','Challenge B']}).to_excel(sheet,index=False)
        files=[('sample.xlsx',sheet.getvalue())]
        self.assertEqual(core.qba_protocol_labels(files),
                         ['Quiet Breathing','Challenge A','Acclimation','Challenge B'])
        with self.assertRaisesRegex(core.AnalysisError, 'Unassigned'):
            core.qba_run(files,prefix='LD',baseline=['Quiet Breathing'],challenge=['Challenge A'],excluded=['Acclimation'])
        with self.assertRaisesRegex(core.AnalysisError, 'exactly one category'):
            core.qba_run(files,prefix='LD',baseline=['Quiet Breathing'],challenge=['Challenge A','Challenge B'],excluded=['Acclimation','Challenge A'])
        text=(core.ASSETS/'QB_Breathing_Tag_Updated_Fin.Rmd').read_text(encoding='utf-8')
        generated=core.make_qba_script(text,Path('/tmp/a.xlsx'),'LD',['Quiet Breathing'],['Acclimation'])
        self.assertIn('challenge_protocols <-',generated)
        self.assertIn('filter(cumulative_time_msec > (end_time_msec - 120000))',generated)
        self.assertIn('if ("3.1" %in% flag_vec) return(3.1)',generated)
        self.assertNotIn('file.choose()',generated)
        self.assertNotIn('readline(',generated)

    def test_disallow_too_large(self):
        with self.assertRaises(core.AnalysisError):
            core.validate_files([('file.xlsx',b'x'*(core.MAX_UPLOAD_BYTES+1))])

if __name__=='__main__': unittest.main()
