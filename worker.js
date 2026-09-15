function unauthorized() {
    return new Response("Authentication required.", {
        status: 401,
        headers: {
            "WWW-Authenticate": 'Basic realm="SKU Finder", charset="UTF-8"',
            "Cache-Control": "no-store"
        }
    });
}

export default {

    async fetch(request, env) {

        const authHeader =
            request.headers.get("Authorization");


        // No login supplied
        if (
            !authHeader ||
            !authHeader.startsWith("Basic ")
        ) {
            return unauthorized();
        }


        try {

            const encoded =
                authHeader.substring(6);

            const decoded =
                atob(encoded);

            const separator =
                decoded.indexOf(":");


            if (separator === -1) {
                return unauthorized();
            }


            const username =
                decoded.substring(0, separator);

            const password =
                decoded.substring(separator + 1);


            // Compare against Cloudflare secrets
            if (
                username !== env.AUTH_USER ||
                password !== env.AUTH_PASS
            ) {

                return unauthorized();

            }


            // Login correct:
            // Allow access to index.html, data.json,
            // CSS, JS, images, etc.
            return env.ASSETS.fetch(request);


        } catch (error) {

            return unauthorized();

        }

    }

};
