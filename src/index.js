export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // Quando o usuário acessar a rota /api/marcas ou a página principal /
    if (url.pathname === "/api/marcas" || url.pathname === "/") {
      try {
        // O garçom (código) pede ao banco de dados (env.DB) para buscar a tabela 'marcas'
        const { results } = await env.DB.prepare("SELECT * FROM marcas").all();

        // Ele devolve os dados em formato de texto para a tela
        return new Response(JSON.stringify(results, null, 2), {
          headers: { 
            "content-type": "application/json;charset=UTF-8",
            "access-control-allow-origin": "*"
          }
        });
      } catch (error) {
        // Se houver algum erro no banco, ele nos avisa o motivo
        return new Response(JSON.stringify({ error: error.message }), {
          status: 500,
          headers: { "content-type": "application/json" }
        });
      }
    }

    return new Response("Página não encontrada", { status: 404 });
  }
};
