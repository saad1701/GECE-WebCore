
import React, { useMemo } from 'react'
import { useForm, Controller } from 'react-hook-form'
import { zodResolver } from '@hookform/resolvers/zod'
import { Grid, Paper, Typography, TextField, Checkbox, FormControlLabel, Button } from '@mui/material'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import api from '../api/client'

function FieldRenderer(props: { name: string; value: any; control: any; path: string }){
  const { name, value, control, path } = props
  const full = path ? path + '.' + name : name
  if (Array.isArray(value)){
    return (<Paper sx={{p:2, mb:2}}>
      <Typography variant="subtitle1">{name}</Typography>
      <Typography variant="body2">Array length: {value.length}</Typography>
    </Paper>)
  }
  if (value !== null && typeof value === 'object'){
    return (
      <Paper sx={{p:2, mb:2}}>
        <Typography variant="subtitle1">{name}</Typography>
        <Grid container spacing={2}>
          {Object.keys(value).map((k)=> (
            <Grid item xs={12} md={6} key={k}>
              <FieldRenderer name={k} value={(value as any)[k]} control={control} path={full} />
            </Grid>
          ))}
        </Grid>
      </Paper>
    )
  }
  if (typeof value === 'boolean'){
    return (
      <FormControlLabel control={<Controller name={full} control={control} render={({ field }) => (<Checkbox {...field} checked={!!field.value} />)} />} label={name} />
    )
  }
  return (
    <Controller name={full} control={control} render={({ field }) => (
      <TextField {...field} fullWidth label={name} type={typeof value === 'number' ? 'number' : 'text'} />
    )} />
  )
}

export default function FormRenderer(props: { tabId: string; title: string }){
  const { tabId, title } = props
  const qc = useQueryClient()
  const q = useQuery({ queryKey:["form", tabId], queryFn: async ()=>{ const r = await api.get('/forms/' + tabId); return r.data } })
  const form = useForm({ resolver: undefined, values: q.data || {} })
  const mut = useMutation({ mutationFn: async (vals:any)=>{ const r = await api.post('/forms/' + tabId, vals); return r.data }, onSuccess: ()=>{ qc.invalidateQueries({ queryKey:["form", tabId] }) } })
  const onSubmit = (vals:any)=>{ mut.mutate(vals) }
  return (
    <div>
      <Typography variant="h5" sx={{mb:2}}>{title}</Typography>
      {q.isLoading ? <Typography>Loading...</Typography> : null}
      {q.isError ? <Typography color="error">Error</Typography> : null}
      {q.data ? (
        <form onSubmit={form.handleSubmit(onSubmit)}>
          <Grid container spacing={2}>
            {Object.keys(q.data).map((k)=> (
              <Grid item xs={12} md={6} key={k}>
                <FieldRenderer name={k} value={q.data[k]} control={form.control} path="" />
              </Grid>
            ))}
          </Grid>
          <div style={{display:'flex', gap:12, marginTop:16}}>
            <Button type="submit" variant="contained">Save</Button>
            <Button type="button" onClick={()=> form.reset(q.data)} disabled={q.isFetching}>Reset</Button>
          </div>
        </form>
      ) : null}
    </div>
  )
}
