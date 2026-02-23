/**
 * Mock WebSocket Server for X21 Web-UI Testing
 *
 * Simulates the Deno backend WebSocket server
 * Provides realistic Claude API streaming responses for testing
 *
 * Automatically finds a free port when port 0 is specified
 */

import { handleWebSocket } from "./websocket-handler.ts";

// Use port 58252 by default; can be overridden with MOCK_SERVER_PORT env var
const PORT = Deno.env.get("MOCK_SERVER_PORT")
  ? parseInt(Deno.env.get("MOCK_SERVER_PORT")!, 10)
  : 8085;

console.log(`🧪 Starting Mock X21 Backend Server...`);

// CORS headers for browser testing
const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
};

let actualPort: number | undefined;

Deno.serve({
  port: PORT,
  handler: (req) => {
    const url = new URL(req.url);

    // Handle CORS preflight
    if (req.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders,
      });
    }

    // Handle WebSocket upgrade requests
    if (url.pathname === "/ws") {
      if (req.headers.get("upgrade") === "websocket") {
        const { socket, response } = Deno.upgradeWebSocket(req);
        handleWebSocket(socket);
        return response;
      }
      return new Response("Expected WebSocket upgrade", { status: 400 });
    }

    // Health check endpoint - includes port information
    if (url.pathname === "/health") {
      const healthData: { status: string; mode: string; port?: number } = {
        status: "ok",
        mode: "mock"
      };
      if (actualPort !== undefined) {
        healthData.port = actualPort;
      }
      return new Response(JSON.stringify(healthData), {
        headers: { "content-type": "application/json", ...corsHeaders },
      });
    }

    // Mock API endpoints for testing
    if (url.pathname === "/api/recent-chats") {
      return new Response(JSON.stringify({ items: [] }), {
        headers: { "content-type": "application/json", ...corsHeaders },
      });
    }

    if (url.pathname === "/api/search-chats") {
      return new Response(JSON.stringify({ items: [] }), {
        headers: { "content-type": "application/json", ...corsHeaders },
      });
    }

    if (url.pathname === "/api/recent-user-messages") {
      return new Response(JSON.stringify({ items: [] }), {
        headers: { "content-type": "application/json", ...corsHeaders },
      });
    }

    if (url.pathname === "/api/messages") {
      return new Response(JSON.stringify({ items: [] }), {
        headers: { "content-type": "application/json", ...corsHeaders },
      });
    }

    return new Response("Mock X21 Backend - Use /ws for WebSocket", {
      status: 404,
      headers: corsHeaders,
    });
  },
  onListen: ({ port }: { port: number }) => {
    actualPort = port;
    console.log(`✅ Mock server listening on ws://localhost:${port}/ws`);
    console.log(`   Health check: http://localhost:${port}/health`);
    // Output port in parseable format for orchestration scripts
    console.log(`MOCK_SERVER_PORT=${port}`);
  },
});
