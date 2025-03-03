class CacheManager {
    constructor(repository) {
        this.cache = [];
        this.repository = repository;
    }

    add(item) {
        this.cache.push({
            ...item,
            timestamp: new Date()
        });
    }

    checkDuplicates(key) {
        const duplicados = new Set();
        const visto = new Set();
        
        for (const item of this.cache) {
            if (item.key === key) {
                if (visto.has(item.key)) {
                    duplicados.add(item);
                } else {
                    visto.add(item.key);
                }
            }
        }
        
        return Array.from(duplicados);
    }

    cleanOldestRow() {
        if (this.cache.length === 0) return;
        
        const oldest = this.cache.reduce((min, current) => 
            current.timestamp < min.timestamp ? current : min
        );
        
        this.cache = this.cache.filter(item => item !== oldest);
    }

    async flush() {
        console.log('Iniciando flush de caché');
        
        for (const item of this.cache) {
            try {
                await this.repository.save(item);
                console.log(`Elemento guardado en BD: ${JSON.stringify(item)}`);
            } catch (error) {
                console.error(`Error al guardar elemento en BD: ${error.message}`);
            }
        }
        
        console.log(`Flush completado. ${this.cache.length} elementos procesados`);
        this.cache = [];
    }

    get size() {
        return this.cache.length;
    }
}

module.exports = CacheManager; 