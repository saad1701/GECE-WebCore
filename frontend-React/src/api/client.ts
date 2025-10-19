import axios from 'axios'
const api = axios.create({ baseURL: '/api/v1' })
api.interceptors.request.use(cfg => {
  const t = localStorage.getItem('access_token')
  if (t) { cfg.headers = cfg.headers || {}; (cfg.headers as any)['Authorization'] = 'Bearer ' + t }
  return cfg
})
export default api
