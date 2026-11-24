import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

interface CreateUserResult {
  email: string;
  status: "created" | "skipped" | "failed";
  reason?: string;
}

Deno.serve(async (req) => {
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    // Verify caller is admin using their JWT
    const authHeader = req.headers.get("Authorization");
    if (!authHeader) {
      return new Response(JSON.stringify({ error: "Unauthorized" }), {
        status: 401,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Create admin client using service role key
    const supabaseAdmin = createClient(
      Deno.env.get("SUPABASE_URL") ?? "",
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "",
      { auth: { autoRefreshToken: false, persistSession: false } }
    );

    // Verify the calling user is an admin
    const supabaseUser = createClient(
      Deno.env.get("SUPABASE_URL") ?? "",
      Deno.env.get("SUPABASE_ANON_KEY") ?? "",
      { global: { headers: { Authorization: authHeader } } }
    );
    const { data: { user }, error: userError } = await supabaseUser.auth.getUser();
    if (userError || !user) {
      return new Response(JSON.stringify({ error: "Unauthorized" }), {
        status: 401,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Check admin status
    const { data: profile } = await supabaseAdmin
      .from("users")
      .select("role")
      .eq("id", user.id)
      .single();
    if (!profile || profile.role !== "admin") {
      return new Response(JSON.stringify({ error: "Admin access required" }), {
        status: 403,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Parse request body
    const { emails, names, password } = await req.json();
    if (!emails || !Array.isArray(emails) || emails.length === 0) {
      return new Response(JSON.stringify({ error: "No emails provided" }), {
        status: 400,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }
    if (!password || password.length < 6) {
      return new Response(JSON.stringify({ error: "Password must be at least 6 characters" }), {
        status: 400,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const results: CreateUserResult[] = [];
    const created: string[] = [];
    const skipped: string[] = [];
    const failed: string[] = [];

    // Process each email
    for (const rawEmail of emails) {
      const email = rawEmail.trim().toLowerCase();
      if (!email || !email.includes("@")) {
        results.push({ email: rawEmail, status: "failed", reason: "Invalid email format" });
        failed.push(email);
        continue;
      }

      try {
        // Check if user already exists
        const { data: existing } = await supabaseAdmin
          .from("users")
          .select("id")
          .eq("email", email)
          .maybeSingle();

        if (existing) {
          results.push({ email, status: "skipped", reason: "Already exists" });
          skipped.push(email);
          continue;
        }

        // Get name for this email if provided
        const name = (names && names[rawEmail.trim()]) || "";

        // Create auth user with full_name in metadata
        const { data: newUser, error: createError } = await supabaseAdmin.auth.admin.createUser({
          email,
          password,
          email_confirm: true,
          user_metadata: { full_name: name },
        });

        if (createError) {
          // Handle duplicate in auth.users (may not be in public.users)
          if (createError.message?.includes("already") || createError.message?.includes("duplicate")) {
            results.push({ email, status: "skipped", reason: "Already exists in auth" });
            skipped.push(email);
          } else {
            results.push({ email, status: "failed", reason: createError.message });
            failed.push(email);
          }
          continue;
        }

        // Update full_name in public.users after trigger creates the row
        if (name) {
          await new Promise(r => setTimeout(r, 300)); // wait for trigger
          await supabaseAdmin.from("users").update({ full_name: name }).eq("id", newUser.user.id);
        }

        // Trigger handle_new_auth_user() auto-inserts into public.users
        results.push({ email, status: "created" });
        created.push(email);

      } catch (e) {
        results.push({ email, status: "failed", reason: String(e) });
        failed.push(email);
      }
    }

    return new Response(
      JSON.stringify({
        summary: {
          total: emails.length,
          created: created.length,
          skipped: skipped.length,
          failed: failed.length,
        },
        results,
      }),
      { headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );

  } catch (e) {
    return new Response(JSON.stringify({ error: String(e) }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
