"""Private local web UI with asynchronous forecasts and live status polling."""
from email.parser import BytesParser
from email.policy import default
import json
import sys
import socket
import logging
import tempfile
import threading
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse,parse_qs
import traceback

from forecast_engine import read_inputs,run_all,write_workbook

BASE=Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
HTML=(BASE/'templates'/'index.html').read_bytes()
CSS=(BASE/'static'/'style.css').read_bytes()
JOBS={}; LOCK=threading.Lock()

def worker(job_id, file_path, horizon, output):
    """Process the Excel workbook away from the request thread and update job progress."""
    def report(done,total,detail):
        with LOCK: JOBS[job_id].update(done=done,total=total,detail=detail,status='running')
    try:
        datasets=read_inputs(file_path)
        total=sum(len(a.columns) for a in datasets.values())
        with LOCK: JOBS[job_id].update(total=total,status='running',detail=f'Loaded {len(datasets)} employee(s)')
        results=run_all(datasets,horizon,report)
        with LOCK: JOBS[job_id].update(detail='Creating Excel workbook and validation charts...')
        write_workbook(results,output)
        summary=[{'employee':e,'months':len(v['history']),'train':next(iter(v['tags'].values()))['train_months'],
                  'validation':next(iter(v['tags'].values()))['validation_months'],'tags':len(v['tags'])} for e,v in results.items()]
        with LOCK: JOBS[job_id].update(status='done',done=total,detail='Forecast completed',summary=summary)
    except Exception as exc:
        with LOCK: JOBS[job_id].update(status='error',detail=str(exc))
        traceback.print_exc()

class Handler(BaseHTTPRequestHandler):
    """Serve the app, receive Excel uploads and expose progress and results endpoints."""
    def response(self,code,data,typ):
        self.send_response(code);self.send_header('Content-Type',typ);self.send_header('Cache-Control','no-store');self.send_header('Content-Length',str(len(data)));self.end_headers();self.wfile.write(data)
    def do_GET(self):
        path=urlparse(self.path).path
        if path=='/':return self.response(200,HTML,'text/html; charset=utf-8')
        if path=='/static/style.css':return self.response(200,CSS,'text/css; charset=utf-8')
        if path.startswith('/api/progress/'):
            jid=path.split('/')[-1]
            with LOCK: job=JOBS.get(jid)
            if not job:return self.response(404,b'{}','application/json')
            data={k:v for k,v in job.items() if k!='file'}
            return self.response(200,json.dumps(data).encode(),'application/json')
        if path.startswith('/download/'):
            jid=path.split('/')[-1]
            with LOCK: job=JOBS.get(jid)
            if not job or job['status']!='done': return self.response(404,b'Not ready','text/plain')
            blob=Path(job['file']).read_bytes()
            self.send_response(200);self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');self.send_header('Content-Disposition','attachment; filename="PhilRice_Forecast_Results.xlsx"');self.send_header('Content-Length',str(len(blob)));self.end_headers();self.wfile.write(blob);return
        return self.response(404,b'Not found','text/plain')
    def do_POST(self):
        if self.path!='/api/start':return self.response(404,b'Not found','text/plain')
        try:
            size=int(self.headers.get('Content-Length','0'))
            if size<1 or size>25*1024*1024: raise ValueError('Please upload a file smaller than 25 MB.')
            data=self.rfile.read(size)
            multipart=BytesParser(policy=default).parsebytes(
                b'Content-Type: '+self.headers['Content-Type'].encode()+b'\r\nMIME-Version: 1.0\r\n\r\n'+data)
            form={part.get_param('name',header='content-disposition'):part for part in multipart.iter_parts()}
            horizon=int(form['horizon'].get_content()) if 'horizon' in form else 12
            if not 1<=horizon<=36:raise ValueError('Forecast months must be between 1 and 36.')
            file=form.get('file')
            filename=file.get_filename() if file else ''
            if not filename or not filename.lower().endswith(('.xlsx','.xlsm')):raise ValueError('Please select an Excel workbook (.xlsx or .xlsm).')
            jid=uuid.uuid4().hex
            output=Path(tempfile.gettempdir())/f'philrice_result_{jid}.xlsx'
            input_file=Path(tempfile.gettempdir())/f'philrice_input_{jid}.xlsx'
            input_file.write_bytes(file.get_payload(decode=True))
            with LOCK: JOBS[jid]={'status':'queued','done':0,'total':0,'detail':'Reading input data...','file':str(output),'summary':[]}
            threading.Thread(target=worker,args=(jid,input_file,horizon,output),daemon=True).start()
            return self.response(200,json.dumps({'job':jid}).encode(),'application/json')
        except Exception as ex:
            return self.response(400,json.dumps({'error':str(ex)}).encode(),'application/json')
    def log_message(self,fmt,*args):pass

def start_app():
    """Start a local-only web server on a free port and open the user's browser."""
    srv = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    port = srv.server_address[1]
    url = f'http://127.0.0.1:{port}'
    print(f'PhilRice local app: {url}', flush=True)
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    srv.serve_forever()

if __name__ == '__main__':
    try:
        start_app()
    except Exception:
        error = traceback.format_exc()
        logfile = Path(tempfile.gettempdir()) / 'PhilRice_Forecast_Error.log'
        logfile.write_text(error, encoding='utf-8')
        if sys.platform.startswith('win'):
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, f'Unable to start PhilRice Forecast.\nLog: {logfile}', 'PhilRice Forecast', 0x10)
        raise
