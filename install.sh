#!/bin/bash                                                                                                        
     ================================================                                                                   
     Academic PPT Skill - 安裝腳本                                                                                      
     ================================================                                                                   
     set -e                                                                                                             
                                                                                                                        
     REPO="Yen0131/academic-ppt"                                                                                        
     RAW_BASE="https://raw.githubusercontent.com/$REPO/main"                                                            
     TMP_DIR=$(mktemp -d)                                                                                               
     trap "rm -rf $TMP_DIR" EXIT                                                                                        
     嘗試判斷 script 目錄                                                                                               
     if [ -f "$(dirname "$0")/SKILL.md" ]; then                                                                         
         SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"                                                                    
     elif [ -f "$HOME/.hermes/skills/productivity/academic-ppt/SKILL.md" ]; then                                        
         SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"                                                                    
     else                                                                                                               
         echo "從 GitHub 下載所需檔案..."                                                                               
         curl -fsSL "$RAW_BASE/SKILL.md" -o "$TMP_DIR/SKILL.md"                                                         
         curl -fsSL "$RAW_BASE/generate_ppt.py" -o "$TMP_DIR/generate_ppt.py"                                           
         SCRIPT_DIR="$TMP_DIR"                                                                                          
     fi                                                                                                                 
                                                                                                                        
     TARGET_DIR="$HOME/.hermes/skills/productivity/academic-ppt"                                                        
                                                                                                                        
     echo "=========================================="                                                                  
     echo "  學術簡報生成器 - 安裝程式"                                                                                 
     echo "=========================================="                                                                  
                                                                                                                        
     FORCE_OVERWRITE=false                                                                                              
     if [ "$1" = "-y" ] || [ "$1" = "--force" ]; then                                                                   
         FORCE_OVERWRITE=true                                                                                           
     fi                                                                                                                 
                                                                                                                        
     if [ -d "$TARGET_DIR" ]; then                                                                                      
         if [ "$FORCE_OVERWRITE" = false ]; then                                                                        
             echo "⚠️  偵測到已安裝，是否覆寫？ (y/N)"                                                                  
             read -r confirm                                                                                            
             if [ "$confirm" != "y" ] && [ "$confirm" != "Y" ]; then                                                    
                 echo "安裝中止"                                                                                        
                 exit 0                                                                                                 
             fi                                                                                                         
         fi                                                                                                             
         rm -rf "$TARGET_DIR"                                                                                           
     fi                                                                                                                 
                                                                                                                        
     echo "📁 建立目錄結構..."                                                                                          
     mkdir -p "$TARGET_DIR/scripts"                                                                                     
                                                                                                                        
     cp "$SCRIPT_DIR/SKILL.md" "$TARGET_DIR/"                                                                           
     cp "$SCRIPT_DIR/generate_ppt.py" "$TARGET_DIR/scripts/"                                                            
                                                                                                                        
     echo "🔍 檢查依賴..."                                                                                              
     if python3 -c "import pptx" 2>/dev/null; then                                                                      
         echo "✓ python-pptx 已安裝"                                                                                    
     else                                                                                                               
         echo "📦 安裝 python-pptx..."                                                                                  
         pip3 install python-pptx -q                                                                                    
     fi                                                                                                                 
                                                                                                                        
     echo ""                                                                                                            
     echo "=========================================="                                                                  
     echo "  安裝完成！"                                                                                                
     echo "=========================================="                                                                  
     echo ""                                                                                                            
     echo "執行方式："                                                                                                  
     echo "  python3 $TARGET_DIR/scripts/generate_ppt.py"
