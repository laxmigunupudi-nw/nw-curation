import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  const url = Deno.env.get('SUPABASE_URL') ?? 'NOT SET'
  const key = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') ?? 'NOT SET'
  
  return new Response(JSON.stringify({
    url_set: url !== 'NOT SET',
    key_set: key !== 'NOT SET',
    key_prefix: key.substring(0, 20),
    method: req.method,
    headers: Object.fromEntries(req.headers.entries())
  }), {
    headers: { ...corsHeaders, 'Content-Type': 'application/json' }
  })
})