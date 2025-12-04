require('dotenv').config();
const OmniHRAuth = require('./auth');

class OmniHRAPIClient {
	constructor() {
		this.auth = new OmniHRAuth();
		this.baseURL = this.auth.baseURL;
	}

	async parseResponse(response) {
		const contentType = response.headers.get('content-type');

		if (response.status === 204 || response.status === 205) {
			return null; // No content
		}

		if (contentType && contentType.includes('application/json')) {
			return await response.json();
		}

		if (contentType && contentType.includes('text/')) {
			return await response.text();
		}

		return response;
	}

	async makeRequest(endpoint, options = {}) {
		try {
			const authHeaders = await this.auth.getAuthHeaders();
			const headers = {
				...authHeaders,
				...options.headers,
			};

			const url = `${this.baseURL}${endpoint}`;

			const response = await fetch(url, {
				...options,
				headers,
			});

			if (response.status === 401) {
				const errorBody = await this.parseResponse(response).catch(() => null);
				throw new Error(
					`Unauthorized (401): Invalid API key or subdomain. ${JSON.stringify(
						errorBody
					)}`
				);
			}

			if (!response.ok) {
				const errorBody = await this.parseResponse(response).catch(() => null);
				throw new Error(
					`HTTP error! status: ${response.status}, body: ${JSON.stringify(
						errorBody
					)}`
				);
			}

			return await this.parseResponse(response);
		} catch (error) {
			this.handleError(error);
			throw error;
		}
	}

	get(endpoint, params = {}) {
		const queryString = new URLSearchParams(params).toString();
		const url = queryString ? `${endpoint}?${queryString}` : endpoint;

		return this.makeRequest(url, {
			method: 'GET',
		});
	}

	async post(endpoint, data = {}) {
		return this.makeRequest(endpoint, {
			method: 'POST',
			body: JSON.stringify(data),
		});
	}

	async put(endpoint, data = {}) {
		return this.makeRequest(endpoint, {
			method: 'PUT',
			body: JSON.stringify(data),
		});
	}

	async delete(endpoint) {
		return this.makeRequest(endpoint, {
			method: 'DELETE',
		});
	}

	handleError(error) {
		if (error.message.includes('HTTP error')) {
			console.error(`API Error: ${error.message}`);
		} else if (error.message.includes('fetch')) {
			console.error('Network Error:', error.message);
		} else {
			console.error('Error:', error.message);
		}
	}
}

module.exports = OmniHRAPIClient;
