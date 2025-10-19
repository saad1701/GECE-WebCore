import React, { createContext, useContext, useEffect, useState } from 'react'
export type User = { id: string; email: string; roles: string[]; tenant_id: string }
const Ctx = createContext<{ user: User | null, setUser: (u: User | null)=>void }>({ user: null, setUser: ()=>{} })
export function AuthProvider({ children }: { children: React.ReactNode }) {
  const [user, setUser] = useState<User | null>(null)
  useEffect(()=>{
    const mock = localStorage.getItem('mock_user')
    if (mock) setUser(JSON.parse(mock))
  },[])
  return <Ctx.Provider value={{user, setUser}}>{children}</Ctx.Provider>
}
export function useAuth(){ return useContext(Ctx) }
