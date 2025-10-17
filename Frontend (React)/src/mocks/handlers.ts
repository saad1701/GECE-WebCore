
import { http, HttpResponse } from 'msw'
import header from './data/data-entry-json.json'
import system from './data/data-entry-json.json'
import device from './data/data-entry-json.json'
import control from './data/data-entry-json.json'
import tmc from './data/data-entry-json.json'
import esd from './data/data-entry-json.json'
import testing from './data/data-entry-json.json'
import documentation from './data/data-entry-json.json'
import meetings from './data/data-entry-json.json'
import report from './data/proposal-summary-json.json'
import freeformat from './data/toolkit-json.json'
import customtraining from './data/toolkit-json.json'
import siteservices from './data/schedule-json.json'
import travelliving from './data/schedule-json.json'
import summary from './data/proposal-summary-json.json'
import metrics from './data/unit-pricing-calculations-json.json'
import about from './data/coversheet-json.json'

const mem:any = {}

function getForm(id:string, src:any){ if(!(id in mem)){ mem[id] = JSON.parse(JSON.stringify(src)) } return mem[id] }
function setForm(id:string, body:any){ mem[id] = body; return mem[id] }

export const handlers = [
  http.get('/forms/header', ()=> HttpResponse.json(getForm('header', header))),
  http.post('/forms/header', async ({ request })=> HttpResponse.json(setForm('header', await request.json()))),

  http.get('/forms/system', ()=> HttpResponse.json(getForm('system', system))),
  http.post('/forms/system', async ({ request })=> HttpResponse.json(setForm('system', await request.json()))),

  http.get('/forms/device-integration', ()=> HttpResponse.json(getForm('device-integration', device))),
  http.post('/forms/device-integration', async ({ request })=> HttpResponse.json(setForm('device-integration', await request.json()))),

  http.get('/forms/control-processor', ()=> HttpResponse.json(getForm('control-processor', control))),
  http.post('/forms/control-processor', async ({ request })=> HttpResponse.json(setForm('control-processor', await request.json()))),

  http.get('/forms/tmc', ()=> HttpResponse.json(getForm('tmc', tmc))),
  http.post('/forms/tmc', async ({ request })=> HttpResponse.json(setForm('tmc', await request.json()))),

  http.get('/forms/esd', ()=> HttpResponse.json(getForm('esd', esd))),
  http.post('/forms/esd', async ({ request })=> HttpResponse.json(setForm('esd', await request.json()))),

  http.get('/forms/testing', ()=> HttpResponse.json(getForm('testing', testing))),
  http.post('/forms/testing', async ({ request })=> HttpResponse.json(setForm('testing', await request.json()))),

  http.get('/forms/documentation', ()=> HttpResponse.json(getForm('documentation', documentation))),
  http.post('/forms/documentation', async ({ request })=> HttpResponse.json(setForm('documentation', await request.json()))),

  http.get('/forms/meetings', ()=> HttpResponse.json(getForm('meetings', meetings))),
  http.post('/forms/meetings', async ({ request })=> HttpResponse.json(setForm('meetings', await request.json()))),

  http.get('/forms/report', ()=> HttpResponse.json(getForm('report', report))),
  http.post('/forms/report', async ({ request })=> HttpResponse.json(setForm('report', await request.json()))),

  http.get('/forms/free-format', ()=> HttpResponse.json(getForm('free-format', freeformat))),
  http.post('/forms/free-format', async ({ request })=> HttpResponse.json(setForm('free-format', await request.json()))),

  http.get('/forms/custom-training', ()=> HttpResponse.json(getForm('custom-training', customtraining))),
  http.post('/forms/custom-training', async ({ request })=> HttpResponse.json(setForm('custom-training', await request.json()))),

  http.get('/forms/site-services', ()=> HttpResponse.json(getForm('site-services', siteservices))),
  http.post('/forms/site-services', async ({ request })=> HttpResponse.json(setForm('site-services', await request.json()))),

  http.get('/forms/travel-living', ()=> HttpResponse.json(getForm('travel-living', travelliving))),
  http.post('/forms/travel-living', async ({ request })=> HttpResponse.json(setForm('travel-living', await request.json()))),

  http.get('/forms/summary', ()=> HttpResponse.json(getForm('summary', summary))),
  http.post('/forms/summary', async ({ request })=> HttpResponse.json(setForm('summary', await request.json()))),

  http.get('/forms/metrics', ()=> HttpResponse.json(getForm('metrics', metrics))),
  http.post('/forms/metrics', async ({ request })=> HttpResponse.json(setForm('metrics', await request.json()))),

  http.get('/forms/about', ()=> HttpResponse.json(getForm('about', about))),
  http.post('/forms/about', async ({ request })=> HttpResponse.json(setForm('about', await request.json()))),
]
