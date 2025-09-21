#!/usr/bin/env node
import { Command } from 'commander'
import path from 'node:path'
import { createRequire } from 'node:module'
import fs from 'fs-extra'
import { parse } from 'jsonc-parser'

const require = createRequire(import.meta.url)
const cliPackage = require('../package.json')

const DEFAULT_LOCATION = 'westus2'
const DEFAULT_BOT_DOMAIN = 'localhost'

async function run() {
  const program = new Command()
  program
    .name('hubot-ms-teams')
    .description('Hubot + Microsoft Teams helper utilities')
    .version(cliPackage.version)

  program
    .command('atk.basic')
    .description('Bootstrap Agent Toolkit (ATK) support in the current project')
    .option('--create-hubot', 'Create a Hubot project if one is not detected')
    .option('--language <language>', 'Preferred project language (ts or js)')
    .option('--subscription <subscription>', 'Azure subscription id to target')
    .option('--resource-group <resourceGroup>', 'Azure resource group name to use')
    .option('--location <location>', 'Azure resource group location', DEFAULT_LOCATION)
    .option('--app-id <appId>', 'Existing Microsoft Bot App Id to reuse')
    .option('--yes', 'Accept defaults without prompting')
    .action(async options => {
      const projectRoot = process.cwd()
      const language = await determineLanguage(projectRoot, options.language)
      const projectPackage = await readPackageJson(projectRoot)
      const projectName = projectPackage?.name ?? path.basename(projectRoot)
      const resourceGroup = options.resourceGroup ?? `${normalizeName(projectName)}-rg`
      const location = options.location ?? DEFAULT_LOCATION
      const subscription = options.subscription ?? ''
      const changes = []

      const hubotDetected = await detectHubot(projectRoot, projectPackage)
      if (!hubotDetected) {
        if (!options.createHubot) {
          console.error('No Hubot project detected. Re-run with --create-hubot to scaffold a new project.')
          process.exitCode = 1
          return
        }
        await scaffoldHubotProject({
          projectRoot,
          projectName,
          language,
          resourceGroup,
          location,
          changes
        })
      }

      await ensureDependencies({
        projectRoot,
        language,
        projectPackage: await readPackageJson(projectRoot),
        changes,
        resourceGroup,
        location
      })

      const ctx = {
        projectRoot,
        projectName,
        language,
        resourceGroup,
        location,
        subscription,
        appId: options.appId ?? '',
        changes
      }

      await generateAtkConfig(ctx)
      await generateAgentToolkitManifest(ctx)
      await ensureSrcEntry(ctx)
      await ensureHelloScript(ctx)
      await ensureVsCodeConfigs(ctx)
      await ensureEnvFiles(ctx)
      await ensureGitIgnore(ctx)
      await ensureManifest(ctx)
      await ensureInfra(ctx)
      await ensureDeployScripts(ctx)

      if (changes.length === 0) {
        console.log('Project already up to date!')
      } else {
        console.log('Agent Toolkit bootstrap complete. Changes:')
        for (const change of changes) {
          console.log(` • ${change}`)
        }
        printNextSteps(ctx)
      }
    })

  await program.parseAsync(process.argv)
}

function normalizeName(name) {
  return name
    .replace(/[^a-zA-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .toLowerCase()
}

async function determineLanguage(projectRoot, override) {
  if (override === 'ts' || override === 'js') {
    return override
  }
  if (override) {
    console.warn(`Unknown language "${override}". Falling back to auto-detection.`)
  }
  if (await fs.pathExists(path.join(projectRoot, 'tsconfig.json'))) {
    return 'ts'
  }
  const pkg = await readPackageJson(projectRoot)
  if (pkg) {
    const deps = { ...pkg.dependencies, ...pkg.devDependencies }
    if (deps?.typescript) {
      return 'ts'
    }
  }
  return 'js'
}

async function detectHubot(projectRoot, pkg) {
  if (!pkg) {
    return fs.pathExists(path.join(projectRoot, 'bin', 'hubot'))
  }
  const deps = {
    ...pkg.dependencies,
    ...pkg.devDependencies,
    ...pkg.peerDependencies
  }
  if (deps && (deps.hubot || deps['@hubot-friends/hubot'])) {
    return true
  }
  if (await fs.pathExists(path.join(projectRoot, 'bin', 'hubot'))) {
    return true
  }
  return false
}

async function readPackageJson(root) {
  const pkgPath = path.join(root, 'package.json')
  if (!(await fs.pathExists(pkgPath))) {
    return null
  }
  try {
    const raw = await fs.readFile(pkgPath, 'utf8')
    if (!raw.trim()) {
      return {}
    }
    return parseJson(raw)
  } catch (err) {
    console.warn(`Failed to read ${pkgPath}:`, err.message)
    return null
  }
}

function parseJson(text) {
  try {
    return JSON.parse(text)
  } catch (err) {
    try {
      return parse(text)
    } catch (error) {
      throw err
    }
  }
}

async function scaffoldHubotProject({ projectRoot, projectName, language, resourceGroup, location, changes }) {
  const pkgPath = path.join(projectRoot, 'package.json')
  if (!(await fs.pathExists(pkgPath))) {
    const pkgTemplate = {
      name: projectName,
      version: '0.1.0',
      private: true,
      type: 'module',
      scripts: {},
      dependencies: {},
      devDependencies: {}
    }
    await writeJsonFile(pkgPath, pkgTemplate)
    changes.push('Created package.json')
  }

  const pkg = (await readPackageJson(projectRoot)) ?? {}
  pkg.dependencies ||= {}
  pkg.devDependencies ||= {}
  pkg.scripts ||= {}
  pkg.type ||= 'module'
  pkg.description ||= 'Hubot + Teams bot with Agent Toolkit'

  const startScript = language === 'ts' ? 'tsx watch src/index.ts' : 'node src/index.js'
  pkg.scripts['build'] ||= language === 'ts' ? 'tsc -p .' : 'node -e "console.log(\'Nothing to build for JavaScript projects.\')"'
  pkg.scripts['start:dev'] ||= startScript
  pkg.scripts['provision:dev'] ||= `az deployment group create --resource-group ${resourceGroup} --template-file infra/main.bicep --parameters @infra/parameters.dev.json${location ? ` --location ${location}` : ''}`
  pkg.scripts['deploy:dev'] ||= 'bash ./scripts/deploy.sh dev'
  pkg.dependencies['hubot'] ||= '^3.4.2'
  pkg.dependencies['dotenv'] ||= '^16.4.5'
  pkg.dependencies['@microsoft/agenttoolkit'] ||= '^0.3.0'
  pkg.dependencies['hubot-ms-teams'] ||= `^${cliPackage.version}`
  if (language === 'ts') {
    pkg.devDependencies['typescript'] ||= '^5.4.5'
    pkg.devDependencies['tsx'] ||= '^4.7.2'
    pkg.devDependencies['@types/node'] ||= '^20.12.7'
  }

  await writeJsonFile(pkgPath, pkg)
  changes.push('Updated package.json for Hubot project')

  const binPath = path.join(projectRoot, 'bin', 'hubot')
  const binContent = "#!/usr/bin/env node\nrequire('hubot/bin/hubot')\n"
  const binChanged = await writeTextFile(binPath, binContent, { makeExecutable: true })
  if (binChanged) {
    changes.push('Created bin/hubot launcher')
  }

  if (language === 'ts') {
    const tsconfigPath = path.join(projectRoot, 'tsconfig.json')
    if (!(await fs.pathExists(tsconfigPath))) {
      const tsconfig = {
        compilerOptions: {
          target: 'ES2022',
          module: 'ESNext',
          moduleResolution: 'node',
          esModuleInterop: true,
          strict: true,
          forceConsistentCasingInFileNames: true,
          skipLibCheck: true,
          outDir: 'dist',
          rootDir: 'src'
        },
        include: ['src/**/*']
      }
      await writeJsonFile(tsconfigPath, tsconfig)
      changes.push('Created tsconfig.json')
    }
  }
}

async function writeJsonFile(filePath, data) {
  const content = JSON.stringify(data, null, 2) + '\n'
  return writeTextFile(filePath, content)
}

async function writeTextFile(filePath, content, options = {}) {
  await fs.ensureDir(path.dirname(filePath))
  if (await fs.pathExists(filePath)) {
    const existing = await fs.readFile(filePath, 'utf8')
    if (existing === content) {
      if (options.makeExecutable) {
        await fs.chmod(filePath, 0o755)
      }
      return false
    }
  }
  await fs.writeFile(filePath, content)
  if (options.makeExecutable) {
    await fs.chmod(filePath, 0o755)
  }
  return true
}

async function ensureDependencies({ projectRoot, language, projectPackage, changes, resourceGroup, location }) {
  if (!projectPackage) {
    return
  }

  const pkgPath = path.join(projectRoot, 'package.json')
  const pkg = {
    ...projectPackage,
    dependencies: { ...(projectPackage.dependencies ?? {}) },
    devDependencies: { ...(projectPackage.devDependencies ?? {}) },
    scripts: { ...(projectPackage.scripts ?? {}) }
  }

  const depsAdded = []
  const ensureDep = (section, name, version) => {
    if (!section[name]) {
      section[name] = version
      depsAdded.push(`${name}@${version}`)
    }
  }

  ensureDep(pkg.dependencies, 'dotenv', '^16.4.5')
  ensureDep(pkg.dependencies, '@microsoft/agenttoolkit', '^0.3.0')
  ensureDep(pkg.dependencies, 'hubot-ms-teams', `^${cliPackage.version}`)

  if (language === 'ts') {
    ensureDep(pkg.devDependencies, 'typescript', '^5.4.5')
    ensureDep(pkg.devDependencies, 'tsx', '^4.7.2')
    ensureDep(pkg.devDependencies, '@types/node', '^20.12.7')
    pkg.scripts['build'] ||= 'tsc -p .'
    pkg.scripts['start:dev'] ||= 'tsx watch src/index.ts'
  } else {
    pkg.scripts['build'] ||= 'node -e "console.log(\'Nothing to build for JavaScript projects.\')"'
    pkg.scripts['start:dev'] ||= 'node src/index.js'
  }

  pkg.scripts['provision:dev'] ||= `az deployment group create --resource-group ${resourceGroup} --template-file infra/main.bicep --parameters @infra/parameters.dev.json${location ? ` --location ${location}` : ''}`
  pkg.scripts['deploy:dev'] ||= 'bash ./scripts/deploy.sh dev'

  const updated = JSON.stringify(pkg)
  const original = JSON.stringify(projectPackage)
  if (depsAdded.length === 0 && updated === original) {
    return
  }

  await writeJsonFile(pkgPath, pkg)
  changes.push('Updated project dependencies and scripts')
}

async function generateAtkConfig({ projectRoot, projectName, language, resourceGroup, location, changes }) {
  const filePath = path.join(projectRoot, 'atk.config.json')
  const config = {
    $schema: 'https://schemas.microsoft.com/teams/agenttoolkit/config.schema.json',
    project: {
      name: projectName,
      language,
      runtime: 'node',
      adapter: 'hubot-ms-teams'
    },
    deploy: {
      resourceGroup,
      location
    },
    endpoints: {
      bot: 'https://REPLACE_WITH_YOUR_BOT.azurewebsites.net',
      messagingEndpoint: '/api/messages'
    },
    capabilities: ['chat', 'mentions', 'cards.adaptive']
  }

  const changed = await writeJsonFile(filePath, config)
  if (changed) {
    changes.push('Wrote atk.config.json')
  }
}

async function generateAgentToolkitManifest({ projectRoot, projectName, changes }) {
  const filePath = path.join(projectRoot, 'agenttoolkit.manifest.json')
  const manifest = {
    $schema: 'https://schemas.microsoft.com/teams/agenttoolkit/manifest.schema.json',
    version: '1.0.0',
    name: projectName,
    description: {
      short: 'Agent Toolkit manifest for Hubot project',
      full: 'Generated by hubot-ms-teams Agent Toolkit bootstrapper.'
    },
    runtime: {
      type: 'node',
      entryPoint: './src/index'
    },
    adapters: ['hubot-ms-teams']
  }

  const changed = await writeJsonFile(filePath, manifest)
  if (changed) {
    changes.push('Wrote agenttoolkit.manifest.json')
  }
}

async function ensureSrcEntry({ projectRoot, language, changes }) {
  const ext = language === 'ts' ? 'ts' : 'js'
  const filePath = path.join(projectRoot, 'src', `index.${ext}`)
  const content = language === 'ts' ? getTypeScriptEntry() : getJavaScriptEntry()
  const changed = await writeTextFile(filePath, content)
  if (changed) {
    changes.push(`Updated src/index.${ext}`)
  }
}

function getTypeScriptEntry() {
  return `import 'dotenv/config'\nimport { Robot } from 'hubot'\nimport teamsAdapter from 'hubot-ms-teams'\nimport { AgentToolkit } from '@microsoft/agenttoolkit'\n\nconst adapter = teamsAdapter({\n  appId: process.env.BOT_APP_ID!,\n  appPassword: process.env.BOT_APP_PASSWORD!,\n  tenantId: process.env.M365_TENANT_ID || undefined\n})\n\nconst toolkit = new AgentToolkit({ configPath: './atk.config.json' })\n\nconst robot = new Robot(adapter, false, 'hubot', false, 'scripts')\ntoolkit.use(robot)\nrobot.loadFile('scripts', 'hello.js')\nrobot.run()\n`
}

function getJavaScriptEntry() {
  return `import 'dotenv/config'\nimport { Robot } from 'hubot'\nimport teamsAdapter from 'hubot-ms-teams'\nimport { AgentToolkit } from '@microsoft/agenttoolkit'\n\nconst adapter = teamsAdapter({\n  appId: process.env.BOT_APP_ID,\n  appPassword: process.env.BOT_APP_PASSWORD,\n  tenantId: process.env.M365_TENANT_ID || undefined\n})\n\nconst toolkit = new AgentToolkit({ configPath: './atk.config.json' })\n\nconst robot = new Robot(adapter, false, 'hubot', false, 'scripts')\ntoolkit.use(robot)\nrobot.loadFile('scripts', 'hello.js')\nrobot.run()\n`
}

async function ensureHelloScript({ projectRoot, changes }) {
  const filePath = path.join(projectRoot, 'scripts', 'hello.js')
  const content = `module.exports = robot => {\n  robot.respond(/ping/i, res => res.send('pong'))\n}\n`
  const changed = await writeTextFile(filePath, content)
  if (changed) {
    changes.push('Created scripts/hello.js')
  }
}

async function ensureVsCodeConfigs({ projectRoot, language, changes }) {
  const launchPath = path.join(projectRoot, '.vscode', 'launch.json')
  const tasksPath = path.join(projectRoot, '.vscode', 'tasks.json')
  const entryFile = language === 'ts' ? 'src/index.ts' : 'src/index.js'

  const launchConfig = {
    version: '0.2.0',
    configurations: [
      {
        type: 'node',
        request: 'launch',
        name: 'Launch Bot (ATK)',
        program: `\${workspaceFolder}/${entryFile}`,
        preLaunchTask: 'Prepare: Provision Azure',
        outFiles: ['${workspaceFolder}/dist/**/*.js'],
        envFile: '${workspaceFolder}/env/.env.local'
      }
    ]
  }

  const tasksConfig = {
    version: '2.0.0',
    tasks: [
      {
        label: 'Build Bot',
        type: 'shell',
        command: 'npm run build'
      },
      {
        label: 'Prepare: Provision Azure',
        type: 'shell',
        command: 'npm run provision:dev'
      },
      {
        label: 'Deploy Bot',
        type: 'shell',
        command: 'npm run deploy:dev'
      },
      {
        label: 'Start Bot (Local)',
        type: 'shell',
        command: 'npm run start:dev',
        presentation: {
          reveal: 'always'
        }
      }
    ]
  }

  const launchChanged = await writeJsonFile(launchPath, launchConfig)
  if (launchChanged) {
    changes.push('Updated .vscode/launch.json')
  }
  const tasksChanged = await writeJsonFile(tasksPath, tasksConfig)
  if (tasksChanged) {
    changes.push('Updated .vscode/tasks.json')
  }
}

async function ensureEnvFiles({ projectRoot, resourceGroup, location, appId, changes }) {
  const localEnvPath = path.join(projectRoot, 'env', '.env.local')
  const cloudEnvPath = path.join(projectRoot, 'env', '.env.cloud')

  const localContent = `# Local development environment variables\nBOT_APP_ID=${appId}\nBOT_APP_PASSWORD=\nM365_TENANT_ID=\n`
  const cloudContent = `# Cloud deployment environment variables\nBOT_APP_ID=${appId}\nBOT_APP_PASSWORD=\nM365_TENANT_ID=\nRESOURCE_GROUP=${resourceGroup}\nAZURE_LOCATION=${location}\n`

  const localChanged = await writeTextFile(localEnvPath, localContent)
  if (localChanged) {
    changes.push('Created env/.env.local')
  }

  const cloudChanged = await writeTextFile(cloudEnvPath, cloudContent)
  if (cloudChanged) {
    changes.push('Created env/.env.cloud')
  }
}

async function ensureGitIgnore({ projectRoot, changes }) {
  const gitignorePath = path.join(projectRoot, '.gitignore')
  let content = ''
  if (await fs.pathExists(gitignorePath)) {
    content = await fs.readFile(gitignorePath, 'utf8')
  }

  if (!content.includes('.env')) {
    const suffix = content.endsWith('\n') || content.length === 0 ? '' : '\n'
    content += `${suffix}# Agent Toolkit environment files\n.env*\n`
    await fs.writeFile(gitignorePath, content)
    changes.push('Updated .gitignore to ignore env files')
  }
}

async function ensureManifest({ projectRoot, appId, changes }) {
  const manifestPath = path.join(projectRoot, 'appPackage', 'manifest.json')
  let manifest
  if (await fs.pathExists(manifestPath)) {
    const raw = await fs.readFile(manifestPath, 'utf8')
    manifest = raw.trim() ? parseJson(raw) : {}
  } else {
    manifest = {}
  }

  manifest.$schema ||= 'https://developer.microsoft.com/json-schemas/teams/v1.16/MicrosoftTeams.schema.json'
  manifest.manifestVersion ||= '1.16'
  manifest.version ||= '1.0.0'
  manifest.id ||= '00000000-0000-0000-0000-000000000000'
  manifest.packageName ||= 'com.example.bot'
  manifest.developer ||= {
    name: 'Your Company',
    websiteUrl: 'https://localhost',
    privacyUrl: 'https://localhost/privacy',
    termsOfUseUrl: 'https://localhost/terms'
  }
  manifest.name ||= {
    short: 'Hubot ATK Bot',
    full: 'Hubot Agent Toolkit Bot'
  }
  manifest.description ||= {
    short: 'Hubot bot enhanced with Agent Toolkit',
    full: 'Generated manifest for Hubot bot running on Microsoft Teams with Agent Toolkit.'
  }
  manifest.validDomains = Array.from(new Set([...(manifest.validDomains ?? []), DEFAULT_BOT_DOMAIN, '127.0.0.1']))
  manifest.bots = manifest.bots ?? [
    {
      botId: appId || '{{BOT_APP_ID}}',
      scopes: ['personal', 'team', 'groupchat'],
      isNotificationOnly: false,
      supportsFiles: false,
      commandLists: []
    }
  ]

  if (manifest.bots.length > 0) {
    manifest.bots[0].botId = appId || manifest.bots[0].botId || '{{BOT_APP_ID}}'
    manifest.bots[0].scopes = Array.from(new Set([...(manifest.bots[0].scopes ?? []), 'personal', 'team', 'groupchat']))
  }

  const changed = await writeJsonFile(manifestPath, manifest)
  if (changed) {
    changes.push('Patched appPackage/manifest.json')
  }
}

async function ensureInfra({ projectRoot, location, changes }) {
  const bicepPath = path.join(projectRoot, 'infra', 'main.bicep')
  const parametersPath = path.join(projectRoot, 'infra', 'parameters.dev.json')

  const bicepLines = [
    'param botAppId string',
    `param location string = '${location}'`,
    '',
    "var sanitizedName = toLower(replace(botAppId, '-', ''))",
    '',
    "resource plan 'Microsoft.Web/serverfarms@2022-09-01' = {",
    "  name: '${sanitizedName}-plan'",
    '  location: location',
    '  sku: {',
    "    name: 'Y1'",
    "    tier: 'Dynamic'",
    '  }',
    '}',
    '',
    "resource site 'Microsoft.Web/sites@2022-09-01' = {",
    "  name: '${sanitizedName}-app'",
    '  location: location',
    "  kind: 'functionapp'",
    '  properties: {',
    '    serverFarmId: plan.id',
    '    httpsOnly: true',
    '  }',
    '}',
    '',
    'output webAppName string = site.name'
  ]
  const bicepContent = `${bicepLines.join('\n')}\n`

  const parametersContent = {
    $schema: 'https://schema.management.azure.com/schemas/2019-04-01/deploymentParameters.json#',
    contentVersion: '1.0.0.0',
    parameters: {
      botAppId: {
        value: '${BOT_APP_ID}'
      },
      location: {
        value: location
      }
    }
  }

  const bicepChanged = await writeTextFile(bicepPath, bicepContent)
  if (bicepChanged) {
    changes.push('Created infra/main.bicep')
  }

  const paramsChanged = await writeJsonFile(parametersPath, parametersContent)
  if (paramsChanged) {
    changes.push('Created infra/parameters.dev.json')
  }
}

async function ensureDeployScripts({ projectRoot, resourceGroup, subscription, changes }) {
  const deployShPath = path.join(projectRoot, 'scripts', 'deploy.sh')
  const deployPsPath = path.join(projectRoot, 'scripts', 'deploy.ps1')
  const azSub = subscription ? ` --subscription ${subscription}` : ''

  const shContent = `#!/usr/bin/env bash\nset -euo pipefail\n\nENV_NAME=\"${'$'}{1:-dev}\"\n\necho \"Deploying Hubot Agent Toolkit bot (environment: ${'$'}ENV_NAME)\"\naz deployment group create --resource-group ${resourceGroup}${azSub} --template-file infra/main.bicep --parameters @infra/parameters.dev.json\n\necho \"Starting web app...\"\naz webapp up --name ${'$'}ENV_NAME-hubot --resource-group ${resourceGroup}${azSub} --runtime \"NODE|20-lts\"\n`

  const psContent = `param(\n  [Parameter(Mandatory = $true)]\n  [string]$EnvironmentName\n)\n\nWrite-Host \"Deploying Hubot Agent Toolkit bot (environment: $EnvironmentName)\"\naz deployment group create --resource-group ${resourceGroup}${azSub} --template-file infra/main.bicep --parameters @infra/parameters.dev.json\n\nWrite-Host \"Starting web app...\"\naz webapp up --name \"$EnvironmentName-hubot\" --resource-group ${resourceGroup}${azSub} --runtime \"NODE|20-lts\"\n`

  const shChanged = await writeTextFile(deployShPath, shContent, { makeExecutable: true })
  if (shChanged) {
    changes.push('Created scripts/deploy.sh')
  }

  const psChanged = await writeTextFile(deployPsPath, psContent)
  if (psChanged) {
    changes.push('Created scripts/deploy.ps1')
  }
}

function printNextSteps({ language }) {
  console.log('\nNext steps:')
  console.log(' • Open VS Code with `code .` and press F5 to start the bot locally.')
  console.log(' • Update env/.env.local with your bot credentials before debugging.')
  console.log(' • Run `npm run deploy:dev` once you are ready to deploy to Azure.')
  if (language === 'ts') {
    console.log(' • Run `npm run build` to produce compiled JavaScript before packaging.')
  }
}

run().catch(err => {
  console.error(err)
  process.exit(1)
})






