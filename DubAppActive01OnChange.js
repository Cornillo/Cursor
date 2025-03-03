function collectlocal(e) {
    const maxRetries = 3;
    let attempt = 0;
    
    const tryCollect = () => {
        attempt++;
        try {
            console.log(`Iniciando intento ${attempt} de recolección de datos...`);
            DubAppCollect.collect(e, "DubAppTotal01");
            console.log('Recolección exitosa');
        } catch (error) {
            // Registrar detalles específicos del error
            console.error(`Error en intento ${attempt}:`, {
                mensaje: error.message,
                codigo: error.code,
                pila: error.stack
            });
            
            if (attempt < maxRetries) {
                const espera = 1000 * attempt; // Incrementa el tiempo de espera con cada intento
                console.log(`Reintentando en ${espera/1000} segundos...`);
                setTimeout(tryCollect, espera);
            } else {
                console.error(`Error fatal después de ${maxRetries} intentos. Por favor, revise los logs anteriores.`);
                throw new Error(`Fallo en la recolección después de ${maxRetries} intentos: ${error.message}`);
            }
        }
    };
    
    tryCollect();
}