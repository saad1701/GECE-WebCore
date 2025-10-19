
import React from 'react'
import { ImageList, ImageListItem, Paper, Typography } from '@mui/material'

const images = [
  { src: '/assets/screens/Detailed Data Entry Form-Header.PNG', title: 'Header' },
  { src: '/assets/screens/Detailed Data Entry Form-System Tab.PNG', title: 'System' },
  { src: '/assets/screens/Detailed Data Entry Form-Device Integration Tab.PNG', title: 'Device Integration' },
  { src: '/assets/screens/Detailed Data Entry Form-Control Processor Tab.PNG', title: 'Control Processor' },
  { src: '/assets/screens/Detailed Data Entry Form-TMC Tab.PNG', title: 'TMC' },
  { src: '/assets/screens/Detailed Data Entry Form-ESD Tab.PNG', title: 'ESD' },
  { src: '/assets/screens/Detailed Data Entry Form-Testing Tab.PNG', title: 'Testing' },
  { src: '/assets/screens/Detailed Data Entry Form-Documentation Tab.PNG', title: 'Documentation' },
  { src: '/assets/screens/Detailed Data Entry Form-Meetings Tab.PNG', title: 'Meetings' },
  { src: '/assets/screens/Detailed Data Entry Form-Report Tab.PNG', title: 'Report' },
  { src: '/assets/screens/Detailed Data Entry Form-Free Format Tab.PNG', title: 'Free Format' },
  { src: '/assets/screens/Detailed Data Entry Form-Custom Training Tab.PNG', title: 'Custom Training' },
  { src: '/assets/screens/Detailed Data Entry Form-Site Services Tab.PNG', title: 'Site Services' },
  { src: '/assets/screens/Detailed Data Entry Form-Travel & Living Tab.PNG', title: 'Travel & Living' },
  { src: '/assets/screens/Detailed Data Entry Form-Summery Tab.PNG', title: 'Summary' },
  { src: '/assets/screens/Detailed Data Entry Form-Metrics Tab.PNG', title: 'Metrics' },
  { src: '/assets/screens/Detailed Data Entry Form-About Tab.PNG', title: 'About' }
]

export default function Screens(){
  return (
    <Paper sx={{p:2}}>
      <Typography variant="h5" sx={{mb:2}}>Screens</Typography>
      <ImageList cols={3} gap={12} sx={{m:0}}>
        {images.map((it)=> (
          <ImageListItem key={it.src}>
            <img src={it.src} alt={it.title} loading="lazy" />
            <Typography variant="subtitle2">{it.title}</Typography>
          </ImageListItem>
        ))}
      </ImageList>
    </Paper>
  )
}
