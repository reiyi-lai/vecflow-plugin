const API_BASE_URL = 'http://localhost:8000';

export interface SummarizeRequest {
  text: string;
}

export interface SummarizeResponse {
  summary: string;
}

export interface CompareRequest {
  clause1: string;
  clause2: string;
}

export interface CompareResponse {
  comparison: string;
  differences: string[];
  recommendations: string[];
}

export interface RedraftRequest {
  text: string;
  instructions?: string;
}

export interface RedraftResponse {
  redraftedText: string;
  changes: string[];
}

export interface AnalyzeRequest {
  text: string;
  prompt: string;
  context?: string;
}

export interface AnalyzeResponse {
  response: string;
  suggestions?: string[];
}

export interface ChatMessage {
  role: 'user' | 'assistant';
  content: string;
  timestamp: Date;
}

class ApiService {
  private async makeRequest<T>(endpoint: string, data: unknown): Promise<T> {
    const response = await fetch(`${API_BASE_URL}${endpoint}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(data),
    });

    if (!response.ok) {
      throw new Error(`API request failed: ${response.statusText}`);
    }

    return response.json();
  }

  async summarize(request: SummarizeRequest): Promise<SummarizeResponse> {
    return this.makeRequest<SummarizeResponse>('/api/summarize', request);
  }

  async compare(request: CompareRequest): Promise<CompareResponse> {
    return this.makeRequest<CompareResponse>('/api/compare', request);
  }

  async redraft(request: RedraftRequest): Promise<RedraftResponse> {
    return this.makeRequest<RedraftResponse>('/api/redraft', request);
  }

  async analyze(request: AnalyzeRequest): Promise<AnalyzeResponse> {
    return this.makeRequest<AnalyzeResponse>('/api/analyze', request);
  }
}

export const apiService = new ApiService();