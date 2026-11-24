import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const authHeader = req.headers.get("Authorization");
    if (!authHeader) {
      return new Response(JSON.stringify({ error: "Unauthorized" }), {
        status: 401, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const supabaseAdmin = createClient(
      Deno.env.get("SUPABASE_URL") ?? "",
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "",
      { auth: { autoRefreshToken: false, persistSession: false } }
    );

    const supabaseUser = createClient(
      Deno.env.get("SUPABASE_URL") ?? "",
      Deno.env.get("SUPABASE_ANON_KEY") ?? "",
      { global: { headers: { Authorization: authHeader } } }
    );
    const { data: { user }, error: userError } = await supabaseUser.auth.getUser();
    if (userError || !user) {
      return new Response(JSON.stringify({ error: "Unauthorized" }), {
        status: 401, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const { contestId, allTaskRows, fetchScore = false } = await req.json();

    if (!contestId) {
      return new Response(JSON.stringify({ error: "contestId required" }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // ── STEP 1: Save ALL task responses ──────────────────────────────────────
    if (allTaskRows && allTaskRows.length > 0) {
      const { error: respError } = await supabaseAdmin
        .from("responses")
        .upsert(allTaskRows, { onConflict: "task_id,field_name" });
      if (respError) {
        console.error("Response upsert error:", respError.message);
      }
    }

    // ── STEP 2: Get all tasks for this user/contest ───────────────────────────
    const { data: allTasks, error: tasksError } = await supabaseAdmin
      .from("tasks")
      .select("id, status")
      .eq("contest_id", contestId)
      .eq("user_id", user.id);

    if (tasksError) {
      return new Response(JSON.stringify({ error: "Failed to fetch tasks: " + tasksError.message }), {
        status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // ── STEP 3: Submit all non-submitted tasks in batches of 50 ──────────────
    const toSubmit = (allTasks || [])
      .filter(t => t.status !== "submitted")
      .map(t => t.id);

    console.log(`Submitting ${toSubmit.length} tasks for user ${user.id}`);

    let submitErrors = 0;
    const batchSize = 50;
    for (let i = 0; i < toSubmit.length; i += batchSize) {
      const batch = toSubmit.slice(i, i + batchSize);
      const { error: submitError } = await supabaseAdmin
        .from("tasks")
        .update({ status: "submitted", submitted_at: new Date().toISOString() })
        .in("id", batch);

      if (submitError) {
        console.error(`Batch ${i/batchSize + 1} submit error:`, submitError.message);
        submitErrors++;
        // Retry this batch once
        await new Promise(r => setTimeout(r, 2000));
        const { error: retryError } = await supabaseAdmin
          .from("tasks")
          .update({ status: "submitted", submitted_at: new Date().toISOString() })
          .in("id", batch);
        if (retryError) {
          console.error(`Batch ${i/batchSize + 1} retry error:`, retryError.message);
        }
      }
    }

    // ── STEP 4: Verify submission ─────────────────────────────────────────────
    const { data: verifyTasks } = await supabaseAdmin
      .from("tasks")
      .select("id, status")
      .eq("contest_id", contestId)
      .eq("user_id", user.id);

    const stillNotSubmitted = (verifyTasks || []).filter(t => t.status !== "submitted").length;
    if (stillNotSubmitted > 0) {
      console.error(`${stillNotSubmitted} tasks still not submitted after update`);
    }

    // ── STEP 5: Fetch score (only if explicitly requested) ───────────────────
    let score = null;
    if (fetchScore) {
      try {
        const scorePromise = supabaseAdmin
          .from("v_user_contest_accuracy")
          .select("*")
          .eq("contest_id", contestId)
          .eq("user_id", user.id)
          .maybeSingle();
        const timeoutPromise = new Promise((_, reject) =>
          setTimeout(() => reject(new Error("score_timeout")), 5000)
        );
        const { data } = await Promise.race([scorePromise, timeoutPromise]);
        score = data || null;
      } catch(e) {
        console.warn("Score query skipped:", e.message);
      }
    }

    return new Response(
      JSON.stringify({
        success: true,
        tasksSubmitted: toSubmit.length,
        stillNotSubmitted,
        score: score || null,
      }),
      { headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );

  } catch (e) {
    console.error("submit-contest error:", e);
    return new Response(JSON.stringify({ error: String(e) }), {
      status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
