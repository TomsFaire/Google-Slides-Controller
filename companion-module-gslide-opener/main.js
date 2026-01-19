const { InstanceBase, runEntrypoint } = require('@companion-module/base')
const UpdateActions = require('./actions.js')

class GoogleSlidesOpenerInstance extends InstanceBase {
	constructor(internal) {
		super(internal)
		console.log('[gslide-opener] Constructor called')
		this.log('info', '=== Constructor called ===')
	}

	async init(config) {
		console.log('[gslide-opener] init() called with config:', JSON.stringify(config))
		this.log('info', '=== init() called ===')
		this.log('info', `Raw config received: ${JSON.stringify(config)}`)
		
		// Merge config with defaults
		this.config = {
			host: '127.0.0.1',
			port: '9595',
			...config
		}

		console.log('[gslide-opener] Merged config:', JSON.stringify(this.config))
		this.log('info', `Merged config: host=${this.config.host}, port=${this.config.port}`)

		// Set initial status
		this.updateStatus('ok', 'Initializing...')
		this.log('info', 'Status set to: Initializing...')

		// Update actions
		this.log('info', 'Calling updateActions()...')
		this.updateActions()
		this.log('info', 'Actions updated')

		// Test connection after module is ready
		this.log('info', 'Starting connection test...')
		this.testConnection().catch(error => {
			this.log('warn', `Connection test failed during init: ${error.message}`)
		})

		this.log('info', '=== Module initialization completed ===')
		console.log('[gslide-opener] init() completed')
	}

	async destroy() {
		console.log('[gslide-opener] destroy() called')
		this.log('info', '=== destroy() called ===')
	}

	async configUpdated(config) {
		console.log('[gslide-opener] configUpdated() called with:', JSON.stringify(config))
		this.log('info', '=== configUpdated() called ===')
		this.log('info', `New config received: ${JSON.stringify(config)}`)
		
		// Merge config with defaults
		this.config = {
			host: '127.0.0.1',
			port: '9595',
			...config
		}
		
		console.log('[gslide-opener] Merged updated config:', JSON.stringify(this.config))
		this.log('info', `Updated merged config: host=${this.config.host}, port=${this.config.port}`)
		await this.testConnection()
	}

	getConfigFields() {
		console.log('[gslide-opener] getConfigFields() called')
		this.log('info', '=== getConfigFields() called ===')
		
		const fields = [
			{
				type: 'static-text',
				id: 'info',
				width: 12,
				label: 'Information',
				value: 'This module controls the Google Slides Opener Electron app. Make sure the app is running on the same computer.',
			},
			{
				type: 'textinput',
				id: 'host',
				label: 'Host',
				width: 6,
				default: '127.0.0.1',
			},
			{
				type: 'number',
				id: 'port',
				label: 'Port',
				width: 6,
				min: 1,
				max: 65535,
				default: 9595,
			},
		]
		
		console.log('[gslide-opener] getConfigFields() returning:', JSON.stringify(fields))
		this.log('info', `Returning ${fields.length} config fields`)
		return fields
	}

	updateActions() {
		console.log('[gslide-opener] updateActions() called')
		this.log('info', '=== updateActions() called ===')
		UpdateActions(this)
		this.log('info', 'Actions updated successfully')
	}

	// Test connection to the API
	async testConnection() {
		console.log('[gslide-opener] testConnection() called')
		this.log('info', '=== Testing connection ===')
		this.log('info', `Connecting to ${this.config.host}:${this.config.port}`)
		try {
			const response = await this.apiRequest('GET', '/api/status')
			console.log('[gslide-opener] Connection successful:', response)
			this.log('info', 'Connected to Google Slides Opener')
			this.log('info', `Response: ${JSON.stringify(response)}`)
			this.updateStatus('ok', 'Connected')
		} catch (error) {
			console.log('[gslide-opener] Connection failed:', error)
			this.log('error', `Failed to connect to Google Slides Opener: ${error.message}`)
			this.updateStatus('error', `Connection failed: ${error.message}`)
		}
	}

	// Make API requests
	async apiRequest(method, endpoint, data = null) {
		const http = require('http')

		console.log(`[gslide-opener] API Request: ${method} ${endpoint}`)
		this.log('debug', `API Request: ${method} ${endpoint}`)

		return new Promise((resolve, reject) => {
			const options = {
				hostname: this.config.host || '127.0.0.1',
				port: this.config.port || 9595,
				path: endpoint,
				method: method,
				headers: {
					'Content-Type': 'application/json',
				},
				timeout: 5000,
			}
			
			console.log('[gslide-opener] Request options:', JSON.stringify(options))

			const req = http.request(options, (res) => {
				let responseData = ''
				res.on('data', (chunk) => {
					responseData += chunk
				})
				res.on('end', () => {
					try {
						const response = JSON.parse(responseData)
						if (res.statusCode === 200) {
							resolve(response)
						} else {
							reject(new Error(response.error || 'Request failed'))
						}
					} catch (error) {
						reject(error)
					}
				})
			})

			req.on('error', (error) => {
				reject(error)
			})

			req.on('timeout', () => {
				req.destroy()
				reject(new Error('Request timeout'))
			})

			if (data) {
				req.write(JSON.stringify(data))
			}
			req.end()
		})
	}
}

console.log('[gslide-opener] Module loaded, calling runEntrypoint...')
runEntrypoint(GoogleSlidesOpenerInstance, [])
console.log('[gslide-opener] runEntrypoint called')
