const { InstanceBase, runEntrypoint } = require('@companion-module/base')
const UpdateActions = require('./actions.js')

class GoogleSlidesOpenerInstance extends InstanceBase {
	constructor(internal) {
		super(internal)
	}

	async init(config) {
		this.log('debug', 'Initializing Google Slides Opener module...')
		this.config = config

		// Set default host/port if not configured
		if (!this.config.host) this.config.host = '127.0.0.1'
		if (!this.config.port) this.config.port = '9595'

		// Set initial status
		this.updateStatus('ok', 'Initializing...')

		// Update actions
		this.updateActions()

		// Test connection after module is ready
		this.testConnection().catch(error => {
			this.log('warn', `Connection test failed during init: ${error.message}`)
		})

		this.log('debug', 'Module initialization completed')
	}

	async destroy() {
		this.log('debug', 'destroy')
	}

	async configUpdated(config) {
		this.log('debug', 'configUpdated')
		this.config = config
		await this.testConnection()
	}

	getConfigFields() {
		return [
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
				type: 'textinput',
				id: 'port',
				label: 'Port',
				width: 6,
				default: '9595',
				regex: '/^\\d+$/',
			},
		]
	}

	updateActions() {
		UpdateActions(this)
	}

	// Test connection to the API
	async testConnection() {
		try {
			const response = await this.apiRequest('GET', '/api/status')
			this.log('info', 'Connected to Google Slides Opener')
			this.updateStatus('ok', 'Connected')
		} catch (error) {
			this.log('error', `Failed to connect to Google Slides Opener: ${error.message}`)
			this.updateStatus('error', `Connection failed: ${error.message}`)
		}
	}

	// Make API requests
	async apiRequest(method, endpoint, data = null) {
		const http = require('http')

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

runEntrypoint(GoogleSlidesOpenerInstance, [])
