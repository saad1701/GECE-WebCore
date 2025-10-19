
import React from 'react'
import ReactDOM from 'react-dom/client'
import { createBrowserRouter, RouterProvider } from 'react-router-dom'
import { QueryClient, QueryClientProvider } from '@tanstack/react-query'
import Splash from './routes/Splash'
import Screens from './routes/Screens'
import Header from './pages/Header'
import System from './pages/System'
import DeviceIntegration from './pages/DeviceIntegration'
import ControlProcessor from './pages/ControlProcessor'
import TMC from './pages/TMC'
import ESD from './pages/ESD'
import Testing from './pages/Testing'
import Documentation from './pages/Documentation'
import Meetings from './pages/Meetings'
import Report from './pages/Report'
import FreeFormat from './pages/FreeFormat'
import CustomTraining from './pages/CustomTraining'
import SiteServices from './pages/SiteServices'
import TravelLiving from './pages/TravelLiving'
import Summary from './pages/Summary'
import Metrics from './pages/Metrics'
import About from './pages/About'
import './mocks/browser'

const qc = new QueryClient()

const router = createBrowserRouter([
  { path:'/splash', element: <Splash /> },
  { path:'/screens', element: <Screens /> },
  { path:'/', element: <Header /> },
  { path:'/system', element: <System /> },
  { path:'/device-integration', element: <DeviceIntegration /> },
  { path:'/control-processor', element: <ControlProcessor /> },
  { path:'/tmc', element: <TMC /> },
  { path:'/esd', element: <ESD /> },
  { path:'/testing', element: <Testing /> },
  { path:'/documentation', element: <Documentation /> },
  { path:'/meetings', element: <Meetings /> },
  { path:'/report', element: <Report /> },
  { path:'/free-format', element: <FreeFormat /> },
  { path:'/custom-training', element: <CustomTraining /> },
  { path:'/site-services', element: <SiteServices /> },
  { path:'/travel-living', element: <TravelLiving /> },
  { path:'/summary', element: <Summary /> },
  { path:'/metrics', element: <Metrics /> },
  { path:'/about', element: <About /> },
])

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <QueryClientProvider client={qc}>
      <RouterProvider router={router} />
    </QueryClientProvider>
  </React.StrictMode>
)
