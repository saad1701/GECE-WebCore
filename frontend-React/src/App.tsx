import React from 'react'
import { Routes, Route, Link, Navigate } from 'react-router-dom'
import ProtectedRoute from './components/ProtectedRoute'
import { AuthProvider, useAuth } from './auth/AuthProvider'
import Header from './pages/Header'
import BestPractices from './pages/BestPractices'
import Controls from './pages/Controls'
import DeviceIntegration from './pages/DeviceIntegration'
import ESD from './pages/ESD'
import TMC from './pages/TMC'
import Testing from './pages/Testing'
import Documentation from './pages/Documentation'
import Meetings from './pages/Meetings'
import Report from './pages/Report'
import FreeFormat from './pages/FreeFormat'
import CustomTraining from './pages/CustomTraining'
import SiteServices from './pages/SiteServices'
import TravelLiving from './pages/TravelLiving'
import SystemPage from './pages/SystemPage'
import Summary from './pages/Summary'
import Metrics from './pages/Metrics'
import About from './pages/About'
function Layout({ children }: { children: React.ReactNode }){
  const tabs = ['header','best-practices','controls','device-integration','esd','tmc','testing','documentation','meetings','report','free-format','custom-training','site-services','travel-living','system','summary','metrics','about']
  const { user, setUser } = useAuth()
  return (<div style={{display:'flex',height:'100vh'}}>
    <nav style={{width:240, borderRight:'1px solid #ddd', padding:12, overflow:'auto'}}>
      <div style={{marginBottom:12}}><strong>GECE</strong></div>
      {tabs.map(t=> <div key={t} style={{margin:'6px 0'}}><Link to={'/app/' + t}>{t.replace('-', ' ')}</Link></div>)}
      <div style={{marginTop:12}}>{user ? <button onClick={()=>{localStorage.removeItem('mock_user'); setUser(null)}}>Logout</button> : null}</div>
    </nav>
    <main style={{flex:1, padding:16, overflow:'auto'}}>{children}</main>
  </div>)
}
function Login(){
  const { setUser } = useAuth()
  return (<div style={{padding:40}}><h2>Login (Mock)</h2><button onClick={()=>{const u={id:'1',email:'user@example.com',roles:['analyst'],tenant_id:'t1'}; localStorage.setItem('mock_user', JSON.stringify(u)); localStorage.setItem('access_token','mock'); setUser(u)}}>Sign In</button></div>)
}
export default function App(){
  return (<AuthProvider>
    <Routes>
      <Route path='/login' element={<Login/>} />
      <Route path='/app/*' element={<ProtectedRoute><Layout>
        <Routes>
          <Route path='' element={<Navigate to='header' replace />} />
          <Route path='header' element={<Header/>} />
          <Route path='best-practices' element={<BestPractices/>} />
          <Route path='controls' element={<Controls/>} />
          <Route path='device-integration' element={<DeviceIntegration/>} />
          <Route path='esd' element={<ESD/>} />
          <Route path='tmc' element={<TMC/>} />
          <Route path='testing' element={<Testing/>} />
          <Route path='documentation' element={<Documentation/>} />
          <Route path='meetings' element={<Meetings/>} />
          <Route path='report' element={<Report/>} />
          <Route path='free-format' element={<FreeFormat/>} />
          <Route path='custom-training' element={<CustomTraining/>} />
          <Route path='site-services' element={<SiteServices/>} />
          <Route path='travel-living' element={<TravelLiving/>} />
          <Route path='system' element={<SystemPage/>} />
          <Route path='summary' element={<Summary/>} />
          <Route path='metrics' element={<Metrics/>} />
          <Route path='about' element={<About/>} />
        </Routes>
      </Layout></ProtectedRoute>} />
      <Route path='*' element={<Navigate to='/app' replace />} />
    </Routes>
  </AuthProvider>)
}
