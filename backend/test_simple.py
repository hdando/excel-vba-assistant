# test_simple.py - Test minimal de Gemini
import os
from dotenv import load_dotenv
import google.generativeai as genai

load_dotenv()

# 1. Vérifier la clé
api_key = os.getenv("GEMINI_API_KEY")
print(f"Clé API : {'✅ Trouvée' if api_key else '❌ Manquante'}")

if api_key:
    # 2. Configurer
    genai.configure(api_key=api_key)
    
    # 3. Tester directement chaque modèle
    models = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro']
    
    for model_name in models:
        print(f"\nTest {model_name}:")
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content("Dis OK")
            print(f"✅ Fonctionne ! Réponse : {response.text.strip()}")
            print(f"\n🎉 UTILISEZ CE MODÈLE : {model_name}")
            break
        except Exception as e:
            print(f"❌ Erreur : {str(e)[:100]}")
else:
    print("\nAjoutez votre clé dans backend/.env :")
    print("GEMINI_API_KEY=votre_clé_ici")