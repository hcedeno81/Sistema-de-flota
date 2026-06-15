import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://pbtrvqaraxotyhtbflnc.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBidHJ2cWFyYXhvdHlodGJmbG5jIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE0NDU1NDgsImV4cCI6MjA5NzAyMTU0OH0.96dkuGbVZ4oyc3Jv9g2dM7vv-W3POD9MZbZCtP2goIM";

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);