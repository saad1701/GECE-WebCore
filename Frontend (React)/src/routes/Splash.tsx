
import React, { useEffect, useState } from 'react'
import { useNavigate } from 'react-router-dom'
import splash from '../assets/splash.png'

export default function Splash(){
  const nav = useNavigate()
  const [fade, setFade] = useState(false)
  useEffect(()=>{
    const t1 = setTimeout(()=> setFade(true), 1400)
    const t2 = setTimeout(()=> nav('/'), 2200)
    return ()=>{ clearTimeout(t1); clearTimeout(t2) }
  }, [nav])
  return (
    <div style={{position:'fixed', inset:0, display:'flex', alignItems:'center', justifyContent:'center', background:'#0b1b2b'}}>
      <img src={splash} alt="splash" style={{width:'80%', maxWidth:720, opacity: fade ? 0 : 1, transition:'opacity 600ms ease'}} />
    </div>
  )
}
