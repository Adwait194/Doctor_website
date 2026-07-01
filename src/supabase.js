import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://zlcojuhenwciudzrrtuj.supabase.co'
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpsY29qdWhlbndjaXVkenJydHVqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg3NTk4MDgsImV4cCI6MjA5NDMzNTgwOH0.sJ0K9Nvr1Ni7lTu3668tWmwZzbJDg21IM_SvOH4caSk'

export const supabase = createClient(supabaseUrl, supabaseKey)
