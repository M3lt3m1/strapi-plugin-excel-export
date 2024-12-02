// https://github.com/vinubabu323/strapi-plugin-excel-export
module.exports = {
    config: {
        "api::product.product": {
            sort: "name",
            columns: [
                "name",
                "brand",
                "description",
                "code",
                "location",
                "weight",
                "size"
            ],
            relations: {
                prices: {
                    columns: ["unit_cost"],
                    relations: {
                        supplier: {
                            columns: ["name"]
                        }
                    }
                },
                category: {
                    columns: ["name"],
                },
                type: {
                    columns: ["name"],
                },
                events: {
                    columns: ["name"],
                },
                certifications: {
                    columns: ["name"],
                },
            },
            locale: "false",
            labels: {
                "name": "Nome",
                "brand": "Marca",
                "description": "Descrizione",
                "code": "Riferimento Interno",
                "weight": "Peso",
                "size": "Dimensione",
                "location": "Posizione",
                "category": "Categoria Prodotto",
                "type": "Tipo Prodotto",
                "events": "Eventi",
                "certifications": "Certificazioni",
                "prices": "Prezzo di vendita"
            }
        }
    }
};