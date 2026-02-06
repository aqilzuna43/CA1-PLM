cat > scripts/migrate.sh << 'EOF'
#!/bin/bash
# Migration helper for refactoring legacy code

echo "📦 CA1-PLM Migration Helper"
echo ""
echo "This script helps you migrate from legacy to modular structure"
echo ""
echo "Steps:"
echo "1. Review legacy files in src/legacy/"
echo "2. Extract reusable functions into src/core/"
echo "3. Update function calls to use new modules"
echo "4. Test thoroughly before removing legacy code"
echo ""
echo "Recommended order:"
echo "  1. Constants.gs - Move all config values"
echo "  2. SheetService.gs - Wrap all getRange() calls"
echo "  3. CacheService.gs - Add caching layer"
echo "  4. HierarchyParser.gs - Extract level parsing logic"
echo "  5. BOMManager.gs - Core BOM operations"
echo "  6. ECRProcessor.gs - ECR workflow"
echo ""
read -p "Press Enter to continue..."
EOF

chmod +x scripts/migrate.sh