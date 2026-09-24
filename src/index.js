export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // Aceita chamadas para /api/marcas ou para a página inicial /
    if (url.pathname === "/api/marcas" || url.pathname === "/") {
      try {
        // Consulta todos os registos guardados na tabela 'marcas' do seu D1
        const { results } = await env.DB.prepare("SELECT * FROM marcas").all();

        return new Response(JSON.stringify(results, null, 2), {
          headers: { 
            "content-type": "application/json;charset=UTF-8",
            "access-control-allow-origin": "*"
          }
        });
      } catch (error) {
        return new Response(JSON.stringify({ error: error.message }), {
          status: 500,
          headers: { "content-type": "application/json" }
        });
      }
    }

    return new Response("Rota não encontrada", { status: 404 });
  }
};
